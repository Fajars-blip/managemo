require('dotenv').config();
const { Telegraf } = require('telegraf');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');
const { GoogleGenAI } = require('@google/genai');

// Inisialisasi Bot, Supabase & Gemini AI
const bot = new Telegraf(process.env.BOT_TOKEN);
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Daftar kategori valid
const KATEGORI_LIST = [
    'Konsumsi',
    'Transportasi',
    'Kebutuhan Rumah',
    'Kesehatan',
    'Pakaian',
    'Elektronik & Pulsa',
    'Hiburan',
    'Pendidikan',
    'Lainnya'
];


// Kata kunci fallback jika AI tidak tersedia
const KEYWORD_MAP = {
    'Konsumsi':          ['makan','minum','kopi','teh','donat','nasi','snack','jajan','bakso','mie','ayam','soto','warteg','warung','resto','restoran','cafe','kafe','pizza','burger','es','juice','minuman','makanan','sarapan','siang','malam','cemilan','roti','kue'],
    'Transportasi':      ['bensin','bbm','parkir','tol','ojek','gojek','grab','bus','angkot','kereta','taksi','taxi','uber','motor','mobil','bensi','bahan bakar','transportasi'],
    'Kebutuhan Rumah':   ['listrik','air','pam','sabun','deterjen','rinso','sunlight','sembako','beras','minyak','gas','lpg','shampoo','pasta gigi','sikat','pel','sapu','tissue','toilet','kebersihan','belanja','supermarket','indomaret','alfamart'],
    'Kesehatan':         ['obat','apotek','dokter','vitamin','masker','rs','rumah sakit','klinik','puskesmas','kesehatan','bpjs','suplemen'],
    'Pakaian':           ['baju','celana','sepatu','kaos','kemeja','rok','jaket','sandal','tas','ikat pinggang','pakaian','baju'],
    'Elektronik & Pulsa':['pulsa','kuota','wifi','internet','charger','kabel','hp','handphone','laptop','earphone','baterai','cas','token'],
    'Hiburan':           ['game','streaming','netflix','spotify','bioskop','film','konser','tiket','wisata','liburan','jalan','hiburan','nongkrong'],
    'Pendidikan':        ['buku','kursus','alat tulis','pensil','pulpen','fotokopi','print','les','sekolah','kuliah','pendidikan','study']
};

// Fungsi klasifikasi kategori menggunakan Gemini AI + keyword fallback
async function getKategori(keterangan) {
    // Coba dengan Gemini AI terlebih dahulu
    if (process.env.GEMINI_API_KEY) {
        try {
            const prompt = `Kamu adalah asisten keuangan Indonesia. Tugasmu mengklasifikasikan 1 item pengeluaran ke dalam kategori yang PALING COCOK.

Kategori yang tersedia (pilih SATU, tulis PERSIS seperti ini):
- Konsumsi (makanan, minuman, kopi, donat, bakso, dll)
- Transportasi (bensin, parkir, ojek, grab, tol, dll)
- Kebutuhan Rumah (listrik, air, sabun, beras, gas, dll)
- Kesehatan (obat, dokter, vitamin, apotek, dll)
- Pakaian (baju, celana, sepatu, dll)
- Elektronik & Pulsa (pulsa, kuota, charger, hp, dll)
- Hiburan (netflix, game, bioskop, dll)
- Pendidikan (buku, kursus, alat tulis, dll)
- Lainnya (jika tidak cocok dengan kategori di atas)

Item pengeluaran: "${keterangan}"

Jawab hanya dengan nama kategorinya saja (satu kata atau frasa):`;

            const response = await ai.models.generateContent({ model: 'gemini-2.5-flash', contents: prompt });
            const result = response.text.trim().replace(/[*_`"']/g, '').trim();
            console.log(`Gemini response for "${keterangan}": "${result}"`);

            // Cari kecocokan persis atau sebagian (case-insensitive)
            const exactMatch = KATEGORI_LIST.find(k => k.toLowerCase() === result.toLowerCase());
            if (exactMatch) return exactMatch;

            const partialMatch = KATEGORI_LIST.find(k => result.toLowerCase().includes(k.toLowerCase()) || k.toLowerCase().includes(result.toLowerCase()));
            if (partialMatch) return partialMatch;
        } catch (err) {
            console.error('Gemini error:', err.message);
        }
    }

    // Fallback: pencocokan kata kunci lokal
    const lower = keterangan.toLowerCase();
    for (const [kategori, keywords] of Object.entries(KEYWORD_MAP)) {
        if (keywords.some(kw => lower.includes(kw))) return kategori;
    }
    return 'Lainnya';
}



// /start - sambutan & panduan penggunaan (tidak menghapus data)
bot.start((ctx) => ctx.reply(
    '👋 Halo! Selamat datang di *ManageMo*.\n\n' +
    '📝 *Cara mencatat pengeluaran:*\n' +
    'Kirim pesan dengan format:\n' +
    '`[Keterangan] [Nominal]`\n' +
    'Contoh: `Makan siang 25000`\n\n' +
    '🤖 AI akan otomatis mengkategorikan pengeluaranmu!\n\n' +
    '📊 *Perintah yang tersedia:*\n' +
    '• `/rekap` — rekap bulan ini\n' +
    '• `/rekap januari` — rekap bulan tertentu\n' +
    '• `/reset` — hapus semua data & mulai dari awal',
    { parse_mode: 'Markdown' }
));


// /reset - alias yang sama dengan konfirmasi
bot.command('reset', async (ctx) => {
    const userId = ctx.from?.id ? ctx.from.id.toString() : null;
    if (!userId) return ctx.reply('❌ Gagal mengidentifikasi user ID Anda.');

    return ctx.reply(
        '⚠️ *Yakin ingin menghapus SEMUA catatan pengeluaran Anda?*\nTindakan ini tidak bisa dibatalkan!',
        {
            parse_mode: 'Markdown',
            reply_markup: {
                inline_keyboard: [[
                    { text: '✅ Ya, hapus semua', callback_data: 'confirm_reset' },
                    { text: '❌ Batal', callback_data: 'cancel_reset' }
                ]]
            }
        }
    );
});

// Handler tombol konfirmasi reset
bot.action('confirm_reset', async (ctx) => {
    const userId = ctx.from?.id ? ctx.from.id.toString() : null;
    await ctx.answerCbQuery();
    if (!userId) return ctx.reply('❌ Gagal mengidentifikasi user ID Anda.');

    const { error } = await supabase
        .from('pengeluaran')
        .delete()
        .eq('user_id', userId);

    if (error) {
        console.error('Error reset:', error);
        return ctx.editMessageText('❌ Gagal menghapus data. Silakan coba lagi.');
    }

    return ctx.editMessageText('🧹 Semua catatan pengeluaran telah dihapus!\n\nSilakan mulai mencatat pengeluaran baru:\nFormat: *[Keterangan] [Nominal]*\nContoh: Makan siang 25000\n\nAI akan mengkategorikan otomatis 🤖', { parse_mode: 'Markdown' });
});

bot.action('cancel_reset', async (ctx) => {
    await ctx.answerCbQuery();
    return ctx.editMessageText('✅ Data Anda aman! Lanjutkan pencatatan pengeluaran.\n\nFormat: *[Keterangan] [Nominal]*\nContoh: Bensin 50000', { parse_mode: 'Markdown' });
});


bot.command('rekap', async (ctx) => {
    const userId = ctx.from?.id ? ctx.from.id.toString() : null;
    if (!userId) return ctx.reply('❌ Gagal mengidentifikasi user ID Anda.');

    const monthsMap = { januari:0,jan:0,februari:1,feb:1,maret:2,mar:2,april:3,apr:3,mei:4,juni:5,jun:5,juli:6,jul:6,agustus:7,agt:7,ags:7,september:8,sep:8,oktober:9,okt:9,november:10,nov:10,desember:11,des:11 };
    const payloadText = ctx.payload || (ctx.message?.text ? ctx.message.text.split(' ').slice(1).join(' ') : '');
    const arg = payloadText.toLowerCase().trim();

    const now = new Date();
    let targetMonth = now.getMonth(), targetYear = now.getFullYear();
    let labelBulan = 'bulan ini';
    let namaBulanDisplay = now.toLocaleString('id-ID', { month: 'long', year: 'numeric' });

    if (arg) {
        if (monthsMap[arg] !== undefined) {
            targetMonth = monthsMap[arg];
            const tempDate = new Date(targetYear, targetMonth, 1);
            namaBulanDisplay = tempDate.toLocaleString('id-ID', { month: 'long', year: 'numeric' });
            labelBulan = `bulan ${namaBulanDisplay}`;
        } else {
            return ctx.reply('❌ Bulan tidak dikenali.\nContoh: /rekap atau /rekap januari');
        }
    }

    ctx.reply(`⏳ Sedang menyusun rekap pengeluaran ${labelBulan}...`);

    const startOfMonth = new Date(targetYear, targetMonth, 1).toISOString();
    const startOfNextMonth = new Date(targetYear, targetMonth + 1, 1).toISOString();

    const { data, error } = await supabase.from('pengeluaran').select('*').eq('user_id', userId).gte('created_at', startOfMonth).lt('created_at', startOfNextMonth).order('created_at', { ascending: true });

    if (error) { console.error("Error tarik data:", error); return ctx.reply('❌ Gagal menarik data dari database.'); }
    if (!data || data.length === 0) return ctx.reply(`📭 Belum ada pengeluaran yang dicatat ${labelBulan === 'bulan ini' ? 'bulan ini' : 'pada ' + labelBulan}.`);

    const workbook = new ExcelJS.Workbook();

    // Sheet 1: Detail Transaksi
    const ws1 = workbook.addWorksheet('Detail Transaksi');
    ws1.columns = [
        { header: 'Tanggal', key: 'tanggal', width: 14 },
        { header: 'Keterangan', key: 'keterangan', width: 30 },
        { header: 'Kategori', key: 'kategori', width: 22 },
        { header: 'Nominal (Rp)', key: 'nominal', width: 16, style: { numFmt: '#,##0' } }
    ];
    const h1 = ws1.getRow(1);
    h1.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    h1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } };
    h1.height = 20;

    let total = 0;
    const kategoriTotals = {};

    data.forEach(row => {
        const d = new Date(row.created_at);
        const kat = row.kategori || 'Lainnya';
        ws1.addRow({ tanggal: `${d.getDate()}/${d.getMonth()+1}/${d.getFullYear()}`, keterangan: row.keterangan, kategori: kat, nominal: row.nominal });
        total += row.nominal;
        kategoriTotals[kat] = (kategoriTotals[kat] || 0) + row.nominal;
    });

    ws1.eachRow((row, rn) => {
        if (rn > 1) row.eachCell(c => { c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: rn%2===0 ? 'FFF1F8E9':'FFFFFFFF' }}; });
    });

    ws1.addRow([]);
    const tr1 = ws1.addRow({ keterangan: 'TOTAL', nominal: total });
    tr1.font = { bold: true };
    tr1.eachCell(c => { c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF9C4' }}; });

    // Sheet 2: Ringkasan Kategori
    const ws2 = workbook.addWorksheet('Ringkasan Kategori');
    ws2.columns = [
        { header: 'Kategori', key: 'kategori', width: 25 },
        { header: 'Total (Rp)', key: 'total', width: 18, style: { numFmt: '#,##0' } },
        { header: 'Persentase', key: 'persen', width: 14, style: { numFmt: '0.0%' } }
    ];
    const h2 = ws2.getRow(1);
    h2.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    h2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1565C0' } };
    h2.height = 20;

    const sorted = Object.entries(kategoriTotals).sort((a,b) => b[1]-a[1]);
    const rowColors = ['FFE3F2FD','FFF3E0','FFE8F5E9','FFFCE4EC','FFF3E5F5','FFFFE0B2','FFE0F7FA','FFF5F5F5','FFEEEEEE'];
    sorted.forEach(([kat, jumlah], idx) => {
        const row = ws2.addRow({ kategori: kat, total: jumlah, persen: jumlah/total });
        row.eachCell(c => { c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: rowColors[idx%rowColors.length] }}; });
    });

    ws2.addRow([]);
    const tr2 = ws2.addRow({ kategori: 'TOTAL', total: total, persen: 1 });
    tr2.font = { bold: true };
    tr2.eachCell(c => { c.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF9C4' }}; });

    const buffer = await workbook.xlsx.writeBuffer();
    const top3 = sorted.slice(0,3).map(([k,v]) => `  • ${k}: Rp${v.toLocaleString('id-ID')}`).join('\n');

    await ctx.replyWithDocument(
        { source: buffer, filename: `Rekap_${namaBulanDisplay.replace(' ','_')}.xlsx` },
        { caption: `📊 Rekap ${labelBulan}\n💰 Total: *Rp${total.toLocaleString('id-ID')}*\n\n🏆 *Pengeluaran terbesar:*\n${top3}`, parse_mode: 'Markdown' }
    );
});

// Handler text
bot.on('text', async (ctx) => {
    const text = ctx.message.text;
    if (text.startsWith('/')) return;

    const match = text.match(/(.+?)\s+(\d+)$/);
    if (!match) return ctx.reply("❌ Format kurang tepat.\nGunakan format: Keterangan [spasi] Angka\nContoh: Makan siang 25000");

    const keterangan = match[1].trim();
    const nominal = parseInt(match[2], 10);
    const userId = ctx.from?.id ? ctx.from.id.toString() : null;
    if (!userId) return ctx.reply("❌ Gagal mengidentifikasi user ID.");

    const kategori = await getKategori(keterangan);

    const { error } = await supabase.from('pengeluaran').insert([{ keterangan, nominal, user_id: userId, kategori }]);
    if (error) { console.error("Error Database:", error); return ctx.reply("❌ Gagal menyimpan data ke database."); }

    ctx.reply(`✅ Tersimpan!\n📝 ${keterangan} — Rp${nominal.toLocaleString('id-ID')}\n🏷️ Kategori: ${kategori}`);
});

// Handler untuk Vercel Serverless Function
module.exports = async (req, res) => {
    try {
        if (req.method === 'POST') { await bot.handleUpdate(req.body); res.status(200).send('OK'); }
        else res.status(200).send('Bot is running...');
    } catch (err) {
        console.error('Error handling update:', err);
        res.status(500).send(err.message || 'Internal Server Error');
    }
};
