require('dotenv').config();
const { Telegraf } = require('telegraf');
const { createClient } = require('@supabase/supabase-js');
const ExcelJS = require('exceljs');

// Inisialisasi Bot & Supabase
const bot = new Telegraf(process.env.BOT_TOKEN);
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

bot.start((ctx) => ctx.reply('Halo! Kirim pengeluaranmu dengan format: [Keterangan] [Nominal]\nContoh: Bensin 20000'));

// Pindahkan bot.command ke atas bot.on('text') agar didahulukan
bot.command('rekap', async (ctx) => {
    ctx.reply('⏳ Sedang menyusun rekap pengeluaran bulan ini...');

    const now = new Date();
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1).toISOString();
    const startOfNextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1).toISOString();

    const { data, error } = await supabase
        .from('pengeluaran')
        .select('*')
        .gte('created_at', startOfMonth)
        .lt('created_at', startOfNextMonth)
        .order('created_at', { ascending: true });

    if (error) {
        console.error("Error tarik data:", error);
        return ctx.reply('❌ Gagal menarik data dari database.');
    }

    if (!data || data.length === 0) {
        return ctx.reply('📭 Belum ada pengeluaran yang dicatat bulan ini.');
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
    const namaBulan = now.toLocaleString('id-ID', { month: 'long', year: 'numeric' });

    await ctx.replyWithDocument(
        { source: buffer, filename: `Rekap_${namaBulan.replace(' ', '_')}.xlsx` },
        { caption: `📊 Total pengeluaranmu bulan ini: *Rp${total.toLocaleString('id-ID')}*`, parse_mode: 'Markdown' }
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

    const { error } = await supabase
        .from('pengeluaran')
        .insert([{ keterangan: keterangan, nominal: nominal }]);

    if (error) {
        console.error("Error Database:", error);
        return ctx.reply("❌ Gagal menyimpan data ke database.");
    }

    ctx.reply(`✅ Tersimpan: ${keterangan} - Rp${nominal.toLocaleString('id-ID')}`);
});

bot.launch();
console.log("Bot pengeluaran sedang berjalan...");

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));