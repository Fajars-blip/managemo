require('dotenv').config();
const { Telegraf } = require('telegraf');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');

// Inisialisasi Bot & Supabase
const bot = new Telegraf(process.env.BOT_TOKEN);
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

bot.start(async (ctx) => {
    const userId = ctx.from?.id ? ctx.from.id.toString() : null;
    if (!userId) {
        return ctx.reply('❌ Gagal mengidentifikasi user ID Anda.');
    }

    const { error } = await supabase
        .from('pengeluaran')
        .delete()
        .eq('user_id', userId);

    if (error) {
        console.error("Error reset data:", error);
        return ctx.reply('❌ Gagal mereset catatan pengeluaran Anda. Silakan coba lagi.');
    }

    return ctx.reply('🧹 Catatan pengeluaran Anda telah dikosongkan (mulai dari awal)!\n\nKirim pengeluaranmu dengan format: [Keterangan] [Nominal]\nContoh: Bensin 20000');
});

// Pindahkan bot.command ke atas bot.on('text') agar didahulukan
bot.command('rekap', async (ctx) => {
    const userId = ctx.from?.id ? ctx.from.id.toString() : null;
    if (!userId) {
        return ctx.reply('❌ Gagal mengidentifikasi user ID Anda.');
    }

    const monthsMap = {
        januari: 0, jan: 0,
        februari: 1, feb: 1,
        maret: 2, mar: 2,
        april: 3, apr: 3,
        mei: 4,
        juni: 5, jun: 5,
        juli: 6, jul: 6,
        agustus: 7, agt: 7, ags: 7,
        september: 8, sep: 8,
        oktober: 9, okt: 9,
        november: 10, nov: 10,
        desember: 11, des: 11
    };

    const payloadText = ctx.payload || (ctx.message?.text ? ctx.message.text.split(' ').slice(1).join(' ') : '');
    const arg = payloadText.toLowerCase().trim();

    const now = new Date();
    let targetMonth = now.getMonth();
    let targetYear = now.getFullYear();
    let namaBulanDisplay = now.toLocaleString('id-ID', { month: 'long', year: 'numeric' });

    if (arg) {
        if (monthsMap[arg] !== undefined) {
            targetMonth = monthsMap[arg];
            const tempDate = new Date(targetYear, targetMonth, 1);
            namaBulanDisplay = tempDate.toLocaleString('id-ID', { month: 'long', year: 'numeric' });
        } else {
            return ctx.reply('❌ Bulan tidak dikenali.\nContoh: `/rekap` (bulan ini) atau `/rekap januari`');
        }
    }

    ctx.reply(`⏳ Sedang menyusun rekap pengeluaran bulan ${namaBulanDisplay}...`);

    const startOfMonth = new Date(targetYear, targetMonth, 1).toISOString();
    const startOfNextMonth = new Date(targetYear, targetMonth + 1, 1).toISOString();

    const { data, error } = await supabase
        .from('pengeluaran')
        .select('*')
        .eq('user_id', userId)
        .gte('created_at', startOfMonth)
        .lt('created_at', startOfNextMonth)
        .order('created_at', { ascending: true });

    if (error) {
        console.error("Error tarik data:", error);
        return ctx.reply('❌ Gagal menarik data dari database.');
    }

    if (!data || data.length === 0) {
        return ctx.reply(`📭 Belum ada pengeluaran yang dicatat pada bulan ${namaBulanDisplay}.`);
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Rekap Pengeluaran');

    // Define columns
    worksheet.columns = [
        { header: 'Tanggal', key: 'tanggal', width: 15 },
        { header: 'Keterangan', key: 'keterangan', width: 30 },
        { header: 'Nominal', key: 'nominal', width: 15, style: { numFmt: '#,##0' } }
    ];

    // Style header row
    worksheet.getRow(1).font = { bold: true };

    let total = 0;

    data.forEach(row => {
        const dateObj = new Date(row.created_at);
        const dateStr = `${dateObj.getDate()}/${dateObj.getMonth() + 1}/${dateObj.getFullYear()}`;
        worksheet.addRow({
            tanggal: dateStr,
            keterangan: row.keterangan,
            nominal: row.nominal
        });
        total += row.nominal;
    });

    // Add empty row
    worksheet.addRow([]);

    // Add total row
    const totalRow = worksheet.addRow({
        keterangan: 'Total',
        nominal: total
    });
    totalRow.getCell('keterangan').font = { bold: true };
    totalRow.getCell('nominal').font = { bold: true };

    const buffer = await workbook.xlsx.writeBuffer();

    await ctx.replyWithDocument(
        { source: buffer, filename: `Rekap_${namaBulanDisplay.replace(' ', '_')}.xlsx` },
        { caption: `📊 Total pengeluaranmu bulan ${namaBulanDisplay}: *Rp${total.toLocaleString('id-ID')}*`, parse_mode: 'Markdown' }
    );
});

// Handler text diletakkan di bawah
bot.on('text', async (ctx) => {
    const text = ctx.message.text;

    // Abaikan jika pesan dimulai dengan '/' (ini perintah command)
    if (text.startsWith('/')) return;

    const match = text.match(/(.+?)\s+(\d+)$/);

    if (!match) {
        return ctx.reply("❌ Format kurang tepat.\nGunakan format: Keterangan [spasi] Angka\nContoh: Makan siang 25000");
    }

    const keterangan = match[1].trim();
    const nominal = parseInt(match[2], 10);
    const userId = ctx.from?.id ? ctx.from.id.toString() : null;

    if (!userId) {
        return ctx.reply("❌ Gagal mengidentifikasi user ID.");
    }

    const { error } = await supabase
        .from('pengeluaran')
        .insert([{ keterangan: keterangan, nominal: nominal, user_id: userId }]);

    if (error) {
        console.error("Error Database:", error);
        return ctx.reply("❌ Gagal menyimpan data ke database.");
    }

    ctx.reply(`✅ Tersimpan: ${keterangan} - Rp${nominal.toLocaleString('id-ID')}`);
});

// Handler untuk Vercel Serverless Function
module.exports = async (req, res) => {
    try {
        if (req.method === 'POST') {
            await bot.handleUpdate(req.body);
            res.status(200).send('OK');
        } else {
            res.status(200).send('Bot is running...');
        }
    } catch (err) {
        console.error('Error handling update:', err);
        res.status(500).send(err.message || 'Internal Server Error');
    }
};
