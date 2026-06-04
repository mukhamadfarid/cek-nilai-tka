let dataPeserta = [];

// Load Excel saat halaman dibuka
window.onload = async () => {
    try {

        const response = await fetch("data.xlsx");
        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const sheetName = workbook.SheetNames[0];

        const worksheet = workbook.Sheets[sheetName];

        dataPeserta = XLSX.utils.sheet_to_json(worksheet);

        console.log(dataPeserta);

    } catch (error) {
        console.error("Gagal membaca file Excel:", error);
    }
};

function cariData() {

    const nisn = document
        .getElementById("nisn")
        .value
        .trim();

    const hasil = document.getElementById("hasil");

    const peserta = dataPeserta.find(
        item => String(item["NISN"]) === nisn
    );

    if (!peserta) {

        hasil.innerHTML = `
            <p class="not-found">
                Data tidak ditemukan
            </p>
        `;

        return;
    }

    hasil.innerHTML = `
        <table>
            <tr>
                <td class="label">NO. PESERTA</td>
                <td>${peserta["NO. PESERTA"] ?? "-"}</td>
            </tr>

            <tr>
                <td class="label">NAMA PESERTA</td>
                <td>${peserta["NAMA PESERTA"] ?? "-"}</td>
            </tr>

            <tr>
                <td class="label">TTL</td>
                <td>${peserta["TTL"] ?? "-"}</td>
            </tr>

            <tr>
                <td class="label">NILAI MATEMATIKA</td>
                <td>${peserta["NILAI MATEMATIKA"] ?? "-"}</td>
            </tr>

            <tr>
                <td class="label">NILAI BAHASA INDONESIA</td>
                <td>${peserta["NILAI BAHASA INDONESIA"] ?? "-"}</td>
            </tr>
        </table>
    `;
}