document.addEventListener('DOMContentLoaded', function() {
    // --- Elemen DOM ---
    const fileInput = document.getElementById('file-input');
    const tabelBody = document.querySelector('#tabel-tagihan tbody');
    const filterCabang = document.getElementById('filter-cabang');
    const filterKategori = document.getElementById('filter-kategori');
    const cariPelanggan = document.getElementById('cari-pelanggan');
    const filterControls = document.getElementById('filter-controls');
    const uploadContainer = document.getElementById('upload-container');
    const clearDataBtn = document.getElementById('clear-data-btn'); // Menambahkan elemen tombol bersihkan data

    // --- Elemen Modal ---
    const lunasModal = document.getElementById('lunas-modal');
    const modalTitle = document.getElementById('modal-title');
    const modalInvoiceList = document.getElementById('modal-invoice-list');
    const closeModalBtn = document.querySelector('.modal-close');
    const prosesLunasBtn = document.getElementById('proses-lunas-btn');

    let dataMaster = [];

    // =====================================================================
    // --- PENYIAPAN EVENT LISTENER ---
    // =====================================================================

    tabelBody.addEventListener('click', function(e) {
        const button = e.target.closest('.btn');

        if (!button) {
            return;
        }

        const pelanggan = button.dataset.pelanggan;

        if (button.classList.contains('btn-wa')) {
            const kategori = button.dataset.kategori;
            kirimPesanWA(pelanggan, kategori);
        }

        if (button.classList.contains('btn-lunas')) {
            bukaModalLunas(pelanggan);
        }
    });

    fileInput.addEventListener('change', handleFile);
    filterCabang.addEventListener('change', filterData);
    filterKategori.addEventListener('change', filterData);
    cariPelanggan.addEventListener('input', filterData);

    prosesLunasBtn.addEventListener('click', () => {
        const pelangganAktif = lunasModal.dataset.currentPelanggan;
        if (pelangganAktif) {
            prosesPembayaran(pelangganAktif);
        }
    });

    closeModalBtn.addEventListener('click', () => lunasModal.style.display = 'none');
    window.addEventListener('click', (e) => {
        if (e.target === lunasModal) lunasModal.style.display = 'none';
    });

    // Event listener untuk tombol bersihkan data
    if (clearDataBtn) {
        clearDataBtn.addEventListener('click', bersihkanData);
    }

    // =====================================================================
    // --- FUNGSI UTAMA & ALUR APLIKASI ---
    // =====================================================================

    function muatDataSaatMulai() {
        const dataTersimpan = localStorage.getItem('dataTagihanTerakhir');
        if (dataTersimpan) {
            try {
                const dataJson = JSON.parse(dataTersimpan);
                if (Array.isArray(dataJson) && dataJson.length > 0) {
                    // Pastikan data diproses ulang untuk normalisasi nomor WA dan validasi angka
                    prosesData(dataJson, false);
                    return;
                }
            } catch (error) {
                console.error("Gagal mem-parsing data dari localStorage:", error);
                localStorage.removeItem('dataTagihanTerakhir');
            }
        }
        tabelBody.innerHTML = '<tr><td colspan="6" style="text-align: center;">Silakan pilih file Excel di atas.</td></tr>';
        filterControls.style.display = 'none'; // Sembunyikan kontrol filter jika tidak ada data
    }

    function handleFile(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                // Menggunakan header: 1 untuk mendapatkan baris pertama sebagai header, lalu mengonversi ke JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'yyyy-mm-dd' });

                // Mendapatkan header dari baris pertama
                const headers = jsonData[0];
                // Memproses sisa data, mengaitkan dengan header
                const actualData = jsonData.slice(1).map(row => {
                    const obj = {};
                    headers.forEach((header, index) => {
                        obj[header] = row[index];
                    });
                    return obj;
                });

                if (actualData.length === 0) {
                    alert("File Excel kosong atau format tidak dapat dibaca.");
                    return;
                }
                prosesData(actualData, true);

            } catch (error) {
                console.error("Terjadi kesalahan saat memproses file Excel:", error);
                alert("Gagal memproses file Excel. Pastikan format file benar dan coba lagi.");
            }
        };
        reader.onerror = function() {
            alert("Tidak dapat membaca file. Silakan coba lagi.");
        };
        reader.readAsArrayBuffer(file);
    }

    function prosesData(data, simpanKeLocalStorage) {
        try {
            dataMaster = data.map(item => {
                // Memastikan kolom-kolom yang diperlukan ada
                const requiredColumns = ['tanggal', 'jatuh tempo', 'pelanggan', 'sisa tagihan', 'no invoice', 'cabang', 'no wa'];
                const missingColumns = requiredColumns.filter(col => item[col] === undefined || item[col] === null);

                if (missingColumns.length > 0) {
                    console.warn(`Data dengan No Invoice '${item['no invoice'] || 'N/A'}' dilewati karena kolom yang hilang: ${missingColumns.join(', ')}.`);
                    return null;
                }

                const tanggal = new Date(item.tanggal);
                const jatuhTempo = new Date(item['jatuh tempo']);

                if (isNaN(tanggal.getTime()) || isNaN(jatuhTempo.getTime())) {
                    console.warn(`Data dengan No Invoice '${item['no invoice']}' memiliki format tanggal yang salah dan dilewati.`);
                    return null;
                }

                // Memastikan 'sisa tagihan' adalah angka
                let sisaTagihanNumerik = parseFloat(item['sisa tagihan']);
                if (isNaN(sisaTagihanNumerik)) {
                    console.warn(`Data dengan No Invoice '${item['no invoice']}' memiliki 'sisa tagihan' yang bukan angka dan dilewati.`);
                    return null;
                }

                let kategori = 'Belum Jatuh Tempo';
                const hariIni = new Date();
                hariIni.setHours(0, 0, 0, 0);
                const tujuhHariSetelahTanggal = new Date(tanggal);
                tujuhHariSetelahTanggal.setDate(tanggal.getDate() + 7);
                const tigaHariSebelumJatuhTempo = new Date(jatuhTempo);
                tigaHariSebelumJatuhTempo.setDate(jatuhTempo.getDate() - 3);

                if (hariIni >= jatuhTempo) kategori = 'Jatuh Tempo';
                else if (hariIni >= tigaHariSebelumJatuhTempo && hariIni < jatuhTempo) kategori = 'Akan Jatuh Tempo';
                else if (hariIni <= tujuhHariSetelahTanggal) kategori = 'Konfirmasi Terima Barang';

                // Normalisasi nomor WhatsApp di sini
                const normalizedNoWA = normalizePhoneNumber(item['no wa']);

                return { ...item, tanggal, 'jatuh tempo': jatuhTempo, 'sisa tagihan': sisaTagihanNumerik, kategori, 'no wa': normalizedNoWA };
            }).filter(item => item !== null);

            if (dataMaster.length === 0) {
                alert("Tidak ada data valid yang dapat diproses dari file.");
                tabelBody.innerHTML = '<tr><td colspan="6" style="text-align: center;">Silakan pilih file Excel di atas.</td></tr>';
                filterControls.style.display = 'none';
                return;
            }

            dataMaster.sort((a, b) => a.pelanggan.localeCompare(b.pelanggan));

            if (simpanKeLocalStorage) {
                try {
                    localStorage.setItem('dataTagihanTerakhir', JSON.stringify(dataMaster));
                } catch (error) {
                    console.error("Gagal menyimpan data ke localStorage:", error);
                    alert("Peringatan: Gagal menyimpan data untuk sesi berikutnya.");
                }
            }

            uploadContainer.querySelector('p').textContent = 'Pilih file Excel lain untuk memperbarui data.';
            filterControls.style.display = 'flex';
            populateFilterCabang();
            filterData();

        } catch (error) {
            console.error("Error pada saat memproses klasifikasi data:", error);
            alert("Terjadi kesalahan saat mengolah data. Periksa kolom-kolom di file Excel Anda.");
        }
    }

    function filterData() {
        const cabang = filterCabang.value;
        const kategori = filterKategori.value;
        const searchTerm = cariPelanggan.value.toLowerCase();
        let filteredData = dataMaster.filter(item => {
            const matchCabang = (cabang === 'semua') || (item.cabang === cabang);
            const matchKategori = (kategori === 'semua') || (item.kategori === kategori);
            const matchPelanggan = item.pelanggan.toLowerCase().includes(searchTerm);
            return matchCabang && matchKategori && matchPelanggan;
        });
        tampilkanData(filteredData);
    }

    function tampilkanData(data) {
        tabelBody.innerHTML = '';
        if (data.length === 0) {
            tabelBody.innerHTML = '<tr><td colspan="6" style="text-align: center;">Tidak ada data yang sesuai.</td></tr>';
            return;
        }
        data.forEach(item => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${item['no invoice']}</td>
                <td>${formatDate(item.tanggal)}</td>
                <td>${formatDate(item['jatuh tempo'])}</td>
                <td>${item.pelanggan}</td>
                <td>${formatRupiah(item['sisa tagihan'])}</td>
                <td class="aksi-container">
                    <button class="btn btn-wa" data-pelanggan="${item.pelanggan}" data-kategori="${item.kategori}">Kirim Pesan</button>
                    <button class="btn btn-lunas" data-pelanggan="${item.pelanggan}">Check Lunas</button>
                </td>
            `;
            tabelBody.appendChild(row);
        });
    }

    function bukaModalLunas(namaPelanggan) {
        modalTitle.textContent = `Pilih Invoice Lunas: ${namaPelanggan}`;
        modalInvoiceList.innerHTML = '';
        lunasModal.dataset.currentPelanggan = namaPelanggan;
        const tagihanPelanggan = dataMaster.filter(item => item.pelanggan === namaPelanggan);
        if (tagihanPelanggan.length === 0) {
            alert('Tidak ada tagihan untuk pelanggan ini.');
            return;
        }
        tagihanPelanggan.forEach(item => {
            const invoiceDiv = document.createElement('div');
            invoiceDiv.className = 'invoice-item';
            invoiceDiv.innerHTML = `
                <input type="checkbox" data-invoice="${item['no invoice']}">
                <div class="invoice-details">
                    <span class="pelanggan-info">${item['no invoice']}</span>
                    <span class="tagihan-info">${formatRupiah(item['sisa tagihan'])}</span>
                </div>
            `;
            modalInvoiceList.appendChild(invoiceDiv);
        });
        lunasModal.style.display = 'flex';
    }

    function prosesPembayaran(namaPelanggan) {
        const checkboxes = modalInvoiceList.querySelectorAll('input[type="checkbox"]:checked');
        if (checkboxes.length === 0) {
            alert('Silakan pilih minimal satu invoice yang sudah dibayar.');
            return;
        }
        const invoiceLunasIds = Array.from(checkboxes).map(cb => cb.dataset.invoice);
        const pelangganData = dataMaster.find(item => item.pelanggan === namaPelanggan);
        if (!pelangganData || !pelangganData['no wa']) {
            alert('Nomor WhatsApp pelanggan tidak ditemukan.');
            return;
        }
        const noWA = pelangganData['no wa'];

        let totalDibayar = 0;
        const invoicesLunasDetail = []; // Menyimpan detail invoice yang lunas
        dataMaster.forEach(item => {
            if (item.pelanggan === namaPelanggan && invoiceLunasIds.includes(item['no invoice'])) {
                totalDibayar += item['sisa tagihan'];
                invoicesLunasDetail.push(item);
            }
        });

        // Filter dataMaster secara global setelah mendapatkan semua invoice yang lunas
        dataMaster = dataMaster.filter(item => !(item.pelanggan === namaPelanggan && invoiceLunasIds.includes(item['no invoice'])));

        let pesan = `Yth. Bapak/Ibu ${namaPelanggan},\n\nTerima kasih atas pembayaran sebesar ${formatRupiah(totalDibayar)}. Pembayaran Anda telah kami terima.\n`;

        // Setelah memfilter dataMaster, hitung sisa tagihan yang benar
        const remainingTagihan = dataMaster.filter(item => item.pelanggan === namaPelanggan);

        if (remainingTagihan.length > 0) {
            let totalSisaTagihan = 0;
            pesan += '\nBerikut adalah sisa tagihan Anda yang masih harus diselesaikan:';
            remainingTagihan.forEach(item => {
                pesan += `\n- ${item['no invoice']} | Jatuh Tempo: ${formatDate(item['jatuh tempo'])} | ${formatRupiah(item['sisa tagihan'])}`;
                totalSisaTagihan += item['sisa tagihan'];
            });
            pesan += `\n\nTotal Sisa Tagihan: ${formatRupiah(totalSisaTagihan)}`;
        } else {
            pesan += '\nSaat ini seluruh tagihan Anda sudah lunas. Terima kasih atas kerjasamanya.';
        }
        pesan += '\n\nSalam hangat.';

        localStorage.setItem('dataTagihanTerakhir', JSON.stringify(dataMaster));
        lunasModal.style.display = 'none';
        filterData();
        const urlWA = `https://web.whatsapp.com/send?phone=${noWA}&text=${encodeURIComponent(pesan)}`;
        window.open(urlWA, 'whatsapp_tab');
    }

    function kirimPesanWA(namaPelanggan, kategori) {
        const tagihanPelanggan = dataMaster.filter(item => item.pelanggan === namaPelanggan);
        if (tagihanPelanggan.length === 0) return;

        const noWA = tagihanPelanggan[0]?.['no wa'];
        if (!noWA) {
            alert('Nomor WhatsApp pelanggan tidak ditemukan.');
            return;
        }

        // Pastikan totalTagihan dihitung dari nilai numerik yang valid
        let totalTagihan = tagihanPelanggan.reduce((sum, item) => {
            const sisaTagihanNum = parseFloat(item['sisa tagihan']);
            return sum + (isNaN(sisaTagihanNum) ? 0 : sisaTagihanNum);
        }, 0);

        // Memastikan formatRupiah() juga digunakan untuk setiap item invoice
        let daftarInvoice = tagihanPelanggan.map(item => {
            const sisaTagihanNum = parseFloat(item['sisa tagihan']);
            const formattedSisaTagihan = isNaN(sisaTagihanNum) ? "Jumlah Tidak Valid" : formatRupiah(sisaTagihanNum);
            return `\n- ${item['no invoice']} (${formattedSisaTagihan})`;
        }).join('');

        let pesan = '';
        const headerPesan = `Yth. Bapak/Ibu ${namaPelanggan},\n\n`;
        // Memastikan totalTagihan diformat dengan formatRupiah()
        const footerPesan = `\n\nTotal Tagihan: ${formatRupiah(totalTagihan)}\n\nTerima kasih.`;

        switch (kategori) {
            case 'Konfirmasi Terima Barang':
                pesan = `${headerPesan}Kami ingin mengkonfirmasi apakah barang untuk tagihan berikut sudah diterima dengan baik?${daftarInvoice}${footerPesan}`;
                break;
            case 'Akan Jatuh Tempo':
                pesan = `${headerPesan}Kami informasikan bahwa tagihan Anda akan segera jatuh tempo. Berikut rinciannya:${daftarInvoice}\n\nMohon untuk dapat segera dilakukan pembayaran.${footerPesan}`;
                break;
            case 'Jatuh Tempo':
                pesan = `${headerPesan}Menurut catatan kami, tagihan berikut telah melewati tanggal jatuh tempo. Berikut rinciannya:${daftarInvoice}\n\nKami mohon untuk segera melakukan pembayaran.${footerPesan}`;
                break;
            default:
                pesan = `${headerPesan}Berikut adalah rincian tagihan Anda:${daftarInvoice}${footerPesan}`;
                break;
        }
        const urlWA = `https://web.whatsapp.com/send?phone=${noWA}&text=${encodeURIComponent(pesan)}`;
        window.open(urlWA, 'whatsapp_tab');
    }

    function populateFilterCabang() {
        const cabangUnik = [...new Set(dataMaster.map(item => item.cabang))];
        filterCabang.innerHTML = '<option value="semua">Semua Cabang</option>';
        cabangUnik.forEach(cabang => {
            const option = document.createElement('option');
            option.value = cabang;
            option.textContent = cabang;
            filterCabang.appendChild(option);
        });
    }

    function formatRupiah(angka) {
        // Tambahkan validasi untuk memastikan angka adalah numerik
        const num = parseFloat(angka);
        if (isNaN(num)) {
            return "Rp0"; // Atau pesan kesalahan lain jika angka tidak valid
        }
        return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(num);
    }

    function formatDate(dateObject) {
        if (!dateObject) return '';
        const date = new Date(dateObject);
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = date.getFullYear();
        return `${day}-${month}-${year}`;
    }

    function normalizePhoneNumber(phoneNumber) {
        if (!phoneNumber) return '';
        let cleaned = String(phoneNumber).replace(/\D/g, ''); // Hapus semua karakter non-digit

        if (cleaned.startsWith('0')) {
            cleaned = '62' + cleaned.substring(1); // Ganti '0' dengan '62'
        } else if (cleaned.startsWith('+')) {
            cleaned = cleaned.substring(1); // Hapus '+'
        }

        // Pastikan dimulai dengan '62'
        if (!cleaned.startsWith('62')) {
            cleaned = '62' + cleaned;
        }
        return cleaned;
    }

    function bersihkanData() {
        if (confirm("Apakah Anda yakin ingin menghapus semua data tagihan? Tindakan ini tidak dapat dibatalkan.")) {
            localStorage.removeItem('dataTagihanTerakhir');
            dataMaster = [];
            tabelBody.innerHTML = '<tr><td colspan="6" style="text-align: center;">Silakan pilih file Excel di atas.</td></tr>';
            uploadContainer.querySelector('p').textContent = 'Pilih file Excel untuk mengunggah data.';
            filterControls.style.display = 'none';
            filterCabang.innerHTML = '<option value="semua">Semua Cabang</option>';
            filterKategori.value = 'semua';
            cariPelanggan.value = '';
            fileInput.value = '';
            alert("Data berhasil dibersihkan.");
        }
    }

    // Terakhir, panggil fungsi untuk memuat data saat aplikasi dimulai
    muatDataSaatMulai();
});