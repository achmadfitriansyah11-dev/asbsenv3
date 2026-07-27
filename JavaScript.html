<script>
        // --- Lazy-load library besar (xlsx, jsPDF) hanya saat dibutuhkan ---
        // PENTING: library ini SENGAJA tidak dimuat otomatis lewat <script src> di head lagi.
        // Salah satu dari library besar ini memicu pelanggaran Content Security Policy sandbox
        // Google Apps Script (CSP memblokir eval()), yang membuat SELURUH script aplikasi ini
        // gagal berjalan sama sekali (halaman jadi blank putih). Dengan dimuat belakangan
        // (baru saat tombol Export Excel/PDF benar-benar diklik), aplikasi utama tetap jalan
        // normal terlepas dari masalah itu.
        const _loadedScripts = {};
        function loadScriptOnce(url) {
            if (_loadedScripts[url]) return _loadedScripts[url];
            _loadedScripts[url] = new Promise((resolve, reject) => {
                const s = document.createElement('script');
                s.src = url;
                s.onload = () => resolve();
                s.onerror = () => { delete _loadedScripts[url]; reject(new Error('Gagal memuat ' + url)); };
                document.head.appendChild(s);
            });
            return _loadedScripts[url];
        }
        let _xlsxLoaded = false, _pdfLoaded = false;
        async function ensureXLSX() {
            if (!_xlsxLoaded) { await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'); _xlsxLoaded = true; }
        }
        async function ensurePDF() {
            if (!_pdfLoaded) {
                await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
                await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js');
                _pdfLoaded = true;
            }
        }

        const MASTER_DB_ID = 'AGENDA_GURU_MASTER_DB';

        // Dummy/Fallback Data for Simulation
        let db = { classes: [], students: [], attendance: [], grades: [], koku: [], journal: [], schedule: [], poin: [], subjects: [], config: {} };
        let chartHadir, chartPrestasi, chartAbsensiObj;
        let editingScheduleId = null;
        let editingJournalId = null;
        let editingKokuId = null;
        let editingPoinId = null;

        function toggleSidebar() {
            document.getElementById('sidebar-menu').classList.toggle('show');
            document.getElementById('sidebar-overlay').classList.toggle('show');
        }

        // --- Core Sync Logic (Diperbarui untuk IndexedDB) ---
        // Antrean simpan: memastikan pengiriman ke Apps Script dilakukan berurutan (tidak tumpang tindih),
        // memberi tahu user jika gagal, dan mencoba lagi otomatis alih-alih diam-diam gagal.
        let syncQueue = Promise.resolve();
        let pendingSyncCount = 0;
        let lastSaveFailed = false;
        // true jika data yang sedang ditampilkan berasal dari cache lokal (perangkat sedang offline),
        // BUKAN dari Spreadsheet — penting supaya user tahu data belum tentu versi terbaru
        // saat berpindah dari perangkat lain.
        let isOfflineMode = false;

        function updateSyncIndicator() {
            const el = document.getElementById('sync-indicator');
            if (!el) return;
            if (pendingSyncCount > 0) {
                el.style.display = 'inline-flex';
                el.innerHTML = '<i class="bi bi-cloud-arrow-up-fill text-primary"></i> <span class="text-primary">Menyimpan...</span>';
            } else if (lastSaveFailed) {
                el.style.display = 'inline-flex';
                el.innerHTML = '<i class="bi bi-cloud-slash-fill text-danger"></i> <span class="text-danger">Gagal sinkron, mencoba lagi...</span>';
            } else if (isOfflineMode) {
                el.style.display = 'inline-flex';
                el.innerHTML = '<i class="bi bi-wifi-off text-warning"></i> <span class="text-warning">Mode Offline (data cache)</span>';
            } else {
                el.style.display = 'none';
            }
        }

        async function syncDB() {
            // Tandai perubahan lokal dengan nomor revisi & waktu, supaya proses sinkronisasi
            // latar belakang bisa mendeteksi apakah data server sudah menyusul atau masih basi.
            db.config._rev = (db.config._rev || 0) + 1;
            db.config._revTime = Date.now();
            const myRev = db.config._rev;
            const dbString = JSON.stringify(db);

            try {
                await localforage.setItem(MASTER_DB_ID, dbString); // Simpan kuat & permanen di database internal aplikasi
                localStorage.setItem(MASTER_DB_ID, dbString); // Fallback cadangan
            } catch(e) { console.error("Gagal menyimpan ke memori lokal", e); }

            // Production Save ke Spreadsheet — diantrekan agar tidak ada dua permintaan simpan
            // yang berjalan bersamaan dan saling menimpa data di server.
            if (typeof google !== 'undefined') {
                pendingSyncCount++; lastSaveFailed = false; updateSyncIndicator();
                syncQueue = syncQueue.then(() => new Promise(resolve => {
                    google.script.run
                        .withSuccessHandler(result => {
                            pendingSyncCount = Math.max(0, pendingSyncCount - 1);
                            lastSaveFailed = false;
                            isOfflineMode = false; // Berhasil kontak server = pasti sedang online
                            // Server adalah pemegang nomor revisi resmi (bisa dinaikkan oleh perangkat lain juga).
                            // Selaraskan revisi lokal dengan revisi resmi dari server agar tidak pernah bentrok
                            // antar-perangkat (HP/Tablet/Laptop) yang memakai akun yang sama.
                            if (result && typeof result.rev !== 'undefined') { db.config._rev = result.rev; }
                            updateSyncIndicator();
                            resolve();
                        })
                        .withFailureHandler(err => {
                            console.error("Gagal menyimpan ke spreadsheet:", err);
                            pendingSyncCount = Math.max(0, pendingSyncCount - 1);
                            lastSaveFailed = true;
                            updateSyncIndicator();
                            // Coba kirim ulang otomatis selama belum ada revisi lain yang lebih baru menggantikannya
                            if (myRev === db.config._rev) {
                                setTimeout(() => syncDB(), 4000);
                            }
                            resolve();
                        })
                        .saveAllData(dbString, MASTER_DB_ID);
                }));
            }
        }

        // Peringatkan user jika masih ada data yang belum selesai disinkronkan ke server
        // saat mereka mencoba menutup/refresh halaman, agar tidak kehilangan data.
        window.addEventListener('beforeunload', (e) => {
            if (pendingSyncCount > 0) {
                e.preventDefault();
                e.returnValue = '';
            }
        });

        // Ambil ulang data langsung dari Spreadsheet dan TIMPA cache lokal browser dengan paksa.
        // Dipakai setelah data di sheet diperbaiki manual (mis. lewat repairLegacyDataTypes di
        // Apps Script), supaya cache lama di perangkat ini tidak balik menimpa data yang sudah dibetulkan.
        function forceRefreshFromServer() {
            if (typeof google === 'undefined') { return Swal.fire('Info', 'Fitur ini hanya berjalan saat aplikasi diakses melalui Google Apps Script.', 'info'); }
            Swal.fire({ title: 'Menyegarkan data...', text: 'Mengambil data terbaru dari Spreadsheet.', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
            google.script.run
                .withSuccessHandler(async res => {
                    if (res && res.config && res.config.registeredEmail) {
                        await localforage.setItem(MASTER_DB_ID, JSON.stringify(res));
                        localStorage.setItem(MASTER_DB_ID, JSON.stringify(res));
                        db = res;
                        closeModal('modalProfile');
                        Swal.fire('Berhasil', 'Data terbaru dari Spreadsheet berhasil dimuat.', 'success').then(() => location.reload());
                    } else {
                        Swal.fire('Gagal', 'Tidak bisa mengambil data dari server.', 'error');
                    }
                })
                .withFailureHandler(err => Swal.fire('Gagal', 'Tidak bisa terhubung ke server: ' + err, 'error'))
                .getAllData(MASTER_DB_ID);
        }

        // --- Initial Load Logic (Server-Authoritative, Aman untuk Multi-Perangkat) ---
        // PENTING: supaya data tidak pernah hilang/tertimpa saat user berpindah dari HP ke
        // tablet lalu ke laptop dengan akun yang sama, aplikasi SELALU mengambil data terbaru
        // dari Spreadsheet lebih dulu setiap kali dibuka — bukan menampilkan cache lokal
        // perangkat itu sendiri. Cache lokal (localforage/localStorage) hanya dipakai sebagai
        // cadangan terakhir kalau perangkat benar-benar tidak bisa menghubungi server (offline),
        // supaya aplikasi tetap bisa dibuka meski datanya mungkin bukan versi terbaru.
        async function getLocalFallbackDb() {
            let localDataStr = await localforage.getItem(MASTER_DB_ID);
            if (!localDataStr) localDataStr = localStorage.getItem(MASTER_DB_ID);
            return localDataStr ? JSON.parse(localDataStr) : null;
        }

        document.addEventListener('DOMContentLoaded', async () => {
            // Safety timeout — diperpanjang karena sekarang selalu menunggu respons server dulu
            setTimeout(() => {
                let loadingEl = document.getElementById('loading-overlay');
                if (loadingEl && loadingEl.style.display !== 'none') {
                    loadingEl.style.display = 'none';
                    if (!localStorage.getItem('agendaGuru_isLogged')) {
                        document.getElementById('login-view').style.display = 'block';
                    }
                }
            }, 15000);

            try { document.getElementById('topbar-date').innerText = new Date().toLocaleDateString('id-ID', {weekday:'long', year:'numeric', month:'long', day:'numeric'}); } catch(e) {}

            document.getElementById('loading-overlay').style.display = 'flex';

            if (typeof google !== 'undefined') {
                // Production Mode (GAS): SELALU tarik data terbaru dari Spreadsheet dulu.
                google.script.run
                    .withSuccessHandler(async res => {
                        isOfflineMode = false;
                        if (res && res.config && res.config.registeredEmail) {
                            await localforage.setItem(MASTER_DB_ID, JSON.stringify(res));
                            localStorage.setItem(MASTER_DB_ID, JSON.stringify(res));
                        }
                        processLoadedData(res);
                    })
                    .withFailureHandler(async err => {
                        // Tidak bisa menghubungi server (mis. sedang offline) -> pakai cache lokal sebagai cadangan
                        console.error("Gagal mengambil data dari server, memakai cache lokal:", err);
                        const localDb = await getLocalFallbackDb();
                        if (localDb && localDb.config && localDb.config.registeredEmail) {
                            isOfflineMode = true;
                            updateSyncIndicator();
                            processLoadedData(localDb);
                        } else {
                            processLoadedData(null);
                        }
                    })
                    .getAllData(MASTER_DB_ID);
            } else {
                // Simulation Mode (Canvas / Browser Local) — tidak ada server GAS untuk dihubungi
                const localDb = await getLocalFallbackDb();
                if (localDb && localDb.config && localDb.config.registeredEmail) {
                    setTimeout(() => processLoadedData(localDb), 300);
                } else {
                    const dummyDB = {
                        config: { registeredEmail: "guru@sekolah.id", registeredPassword: "123", userName: "Ahmad Guru, M.Pd.", userSchool: "SDN 1 Dummy" },
                        classes: [{ ClassID: "C1", ClassName: "Kelas 5A" }, { ClassID: "C2", ClassName: "Kelas 5B" }],
                        students: [
                            { StudentID: "S1", ClassID: "C1", StudentName: "Andi Saputra", ParentPhone: "08123456789" },
                            { StudentID: "S2", ClassID: "C1", StudentName: "Budi Santoso", ParentPhone: "08129876543" },
                            { StudentID: "S3", ClassID: "C1", StudentName: "Citra Kirana", ParentPhone: "08120000111" }
                        ],
                        subjects: ["Matematika", "IPA", "Bahasa Indonesia"],
                        attendance: [], grades: [], koku: [], journal: [], schedule: [], poin: []
                    };
                    let dummyStr = JSON.stringify(dummyDB);
                    await localforage.setItem(MASTER_DB_ID, dummyStr);
                    localStorage.setItem(MASTER_DB_ID, dummyStr);
                    setTimeout(() => processLoadedData(dummyDB), 500);
                }
            }
        });

        function processLoadedData(res) {
            document.getElementById('loading-overlay').style.display = 'none';
            if (res && res.config && res.config.registeredEmail) {
                db = res;
                db.subjects = db.subjects || [];
                if(!db.schedule) db.schedule = [];
                if(!db.poin) db.poin = [];
                
                if (localStorage.getItem('agendaGuru_isLogged') === 'true') {
                    showApp();
                } else {
                    document.getElementById('login-view').style.display = 'block';
                    document.getElementById('form-login').style.display = 'block';
                    document.getElementById('form-register').style.display = 'none';
                }
            } else {
                db = { classes: [], students: [], attendance: [], grades: [], koku: [], journal: [], schedule: [], poin: [], subjects: [], config: {} };
                document.getElementById('login-view').style.display = 'block';
                document.getElementById('form-login').style.display = 'none';
                document.getElementById('form-register').style.display = 'block';
            }
        }

        // --- Auth Logic ---
        function handleAuth(e, mode) {
            e.preventDefault();
            if (mode === 'login') {
                let email = document.getElementById('loginUser').value.trim();
                let pass = document.getElementById('loginPass').value;
                if (!db.config.registeredEmail) return Swal.fire('Akses Ditolak', 'Belum ada akun terdaftar.', 'warning');
                if (email !== db.config.registeredEmail || pass !== db.config.registeredPassword) return Swal.fire('Gagal Masuk', 'Email atau Password salah.', 'error');
                localStorage.setItem('agendaGuru_isLogged', 'true');
                showApp();
            } else if (mode === 'register') {
                let name = document.getElementById('regName').value.trim();
                let email = document.getElementById('regEmail').value.trim();
                let pass = document.getElementById('regPass').value;
                if (db.config.registeredEmail) return Swal.fire('Pendaftaran Ditolak', 'Aplikasi ini sudah memiliki pengguna.', 'error');
                
                db.config.registeredEmail = email; db.config.registeredPassword = pass; db.config.userName = name; db.config.userSchool = "Sekolah Saya";
                syncDB(); localStorage.setItem('agendaGuru_isLogged', 'true');
                Swal.fire({icon: 'success', title: 'Berhasil', text: 'Akun berhasil dibuat!', timer: 1500, showConfirmButton: false});
                showApp();
            }
        }

        function showApp() {
            document.getElementById('login-view').style.display = 'none';
            document.getElementById('app-view').style.display = 'block';
            populateClassSelects();
            populateMapelSelects();
            updateProfileUI();
            switchView('beranda');
        }

        function toggleAuthMode(mode) {
            if (mode === 'register') {
                document.getElementById('form-login').style.display = 'none'; document.getElementById('form-register').style.display = 'block';
            } else {
                document.getElementById('form-login').style.display = 'block'; document.getElementById('form-register').style.display = 'none';
            }
        }

        function logout() { 
            localStorage.removeItem('agendaGuru_isLogged');
            document.getElementById('app-view').style.display = 'none'; document.getElementById('login-view').style.display = 'block';
            document.getElementById('form-login').style.display = 'block'; document.getElementById('form-register').style.display = 'none';
            document.getElementById('loginUser').value = ''; document.getElementById('loginPass').value = '';
            
            // Mereset UI kembali ke beranda tanpa melakukan reload paksa
            switchView('beranda'); 
        }

        // --- Navigation ---
        function switchView(viewId) {
            document.querySelectorAll('.view-section').forEach(el => el.classList.remove('active'));
            const targetView = document.getElementById('view-' + viewId);
            if(targetView) targetView.classList.add('active');
            
            document.querySelectorAll('.nav-link').forEach(el => el.classList.remove('active'));
            const activeLink = document.querySelector(`.nav-link[onclick*="switchView('${viewId}')"]`);
            if (activeLink) activeLink.classList.add('active');
            
            let titles = {beranda: "Beranda", kelas: "Manajemen Kelas", absensi: "Absensi Siswa", nilai: "Data Nilai", kokurikuler: "Nilai Kokurikuler", jurnal: "Jurnal Harian", jadwal: "Jadwal Pelajaran", poin: "Poin Siswa"};
            document.getElementById('topbar-title').innerText = titles[viewId];

            if(viewId === 'beranda') renderBeranda();
            if(viewId === 'absensi') loadAbsensiData();
            if(viewId === 'nilai') loadNilaiData();
            if(viewId === 'kokurikuler') { renderKokuDimensi(); loadKokuHistory(); }
            if(viewId === 'jurnal') { renderJurnal(); document.getElementById('jurnal-date').value = new Date().toISOString().split('T')[0]; }
            if(viewId === 'jadwal') renderSchedule();
            if(viewId === 'poin') { loadPoinData(); document.getElementById('poin-tanggal').value = new Date().toISOString().split('T')[0]; }

            if(window.innerWidth < 992) { document.getElementById('sidebar-menu').classList.remove('show'); document.getElementById('sidebar-overlay').classList.remove('show'); }
        }

        // --- Data Population Helpers ---
        function populateClassSelects() {
            // Untuk pilihan Manajemen Kelas yang hanya menampilkan kelas spesifik (tanpa "Semua Kelas")
            let optsManajemen = '<option value="">-- Pilih Kelas --</option>';
            db.classes.forEach(c => optsManajemen += `<option value="${c.ClassID}">${c.ClassName}</option>`);
            const elManajemen = document.getElementById('select-kelas-manajemen');
            if(elManajemen) {
                let currentVal = elManajemen.value;
                elManajemen.innerHTML = optsManajemen;
                if(currentVal && currentVal !== 'All') elManajemen.value = currentVal;
            }

            // Untuk pilihan filter di menu lainnya (dengan tambahan opsi "Semua Kelas")
            const selects = ['absen-kelas', 'nilai-kelas', 'koku-kelas', 'sched-class', 'filter-sched-class', 'jurnal-kelas', 'filter-jurnal-kelas', 'poin-kelas', 'filter-poin-kelas'];
            let optsAll = '<option value="">-- Pilih Kelas --</option><option value="All">Semua Kelas</option>';
            db.classes.forEach(c => optsAll += `<option value="${c.ClassID}">${c.ClassName}</option>`);
            selects.forEach(id => { 
                let el = document.getElementById(id);
                if(el) {
                    let val = el.value;
                    el.innerHTML = optsAll; 
                    if(val) el.value = val;
                }
            });
        }

        function populateMapelSelects() {
            const selects = ['nilai-mapel', 'jurnal-mapel', 'sched-subject'];
            let opts = '<option value="">-- Pilih Mapel --</option>';
            if(db.subjects) {
                db.subjects.forEach(m => opts += `<option value="${m}">${m}</option>`);
            }
            selects.forEach(id => { 
                let el = document.getElementById(id);
                if(el) {
                    let val = el.value;
                    el.innerHTML = opts; 
                    if(val && db.subjects.includes(val)) el.value = val;
                }
            });
        }

        function showAddMapelModal() {
            Swal.fire({
                title: 'Tambah Mata Pelajaran', input: 'text', inputPlaceholder: 'Cth: Matematika / IPAS',
                showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed && result.value) {
                    let newMapel = result.value.trim();
                    if(!db.subjects) db.subjects = [];
                    if(!db.subjects.includes(newMapel)) {
                        db.subjects.push(newMapel); populateMapelSelects(); syncDB();
                        Swal.fire('Tersimpan!', `Mata Pelajaran ${newMapel} berhasil ditambahkan.`, 'success');
                    } else { Swal.fire('Info', 'Mata Pelajaran tersebut sudah ada di daftar.', 'info'); }
                }
            });
        }

        function escapeHtml(str) {
            return String(str).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
        }

        function showManageMapelModal() {
            if (!db.subjects) db.subjects = [];
            let listHtml;
            if (db.subjects.length === 0) {
                listHtml = '<p class="text-muted text-center mb-0">Belum ada mata pelajaran.</p>';
            } else {
                listHtml = '<ul class="list-group text-start">' + db.subjects.map(m => `
                    <li class="list-group-item d-flex justify-content-between align-items-center">
                        <span>${escapeHtml(m)}</span>
                        <span class="text-nowrap">
                            <button class="btn btn-sm btn-outline-primary me-1" onclick="renameMapel('${encodeURIComponent(m)}')" title="Edit Nama"><i class="bi bi-pencil-square"></i></button>
                            <button class="btn btn-sm btn-outline-danger" onclick="deleteMapel('${encodeURIComponent(m)}')" title="Hapus"><i class="bi bi-trash"></i></button>
                        </span>
                    </li>`).join('') + '</ul>';
            }
            Swal.fire({
                title: 'Kelola Mata Pelajaran',
                html: `<div style="max-height:320px; overflow-y:auto;">${listHtml}</div>`,
                showCancelButton: false,
                confirmButtonText: 'Tutup'
            });
        }

        function renameMapel(encodedOldName) {
            const oldName = decodeURIComponent(encodedOldName);
            Swal.fire({
                title: 'Edit Nama Mata Pelajaran', input: 'text', inputValue: oldName,
                showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal'
            }).then((result) => {
                if (!result.isConfirmed || !result.value) return;
                const newName = result.value.trim();
                if (!newName || newName === oldName) return showManageMapelModal();
                if (db.subjects.includes(newName)) { Swal.fire('Gagal', 'Nama mata pelajaran tersebut sudah ada.', 'error').then(showManageMapelModal); return; }

                const idx = db.subjects.indexOf(oldName);
                if (idx > -1) db.subjects[idx] = newName;

                // Perbarui juga seluruh data Nilai, Jurnal, dan Jadwal yang sudah memakai nama lama,
                // supaya riwayat data yang sudah diinput tidak terlepas dari mata pelajaran barunya.
                db.grades.forEach(g => { if (g.Subject === oldName) g.Subject = newName; });
                db.journal.forEach(j => { if (j.Subject === oldName) j.Subject = newName; });
                db.schedule.forEach(s => { if (s.Subject === oldName) s.Subject = newName; });

                populateMapelSelects(); syncDB();

                // Render ulang tampilan yang mungkin sedang menampilkan data mapel ini
                if (document.getElementById('view-nilai').classList.contains('active')) loadNilaiData();
                if (document.getElementById('view-jurnal').classList.contains('active')) renderJurnal();
                if (document.getElementById('view-jadwal').classList.contains('active')) renderSchedule();

                Swal.fire('Tersimpan!', `Mata Pelajaran berhasil diubah menjadi "${newName}".`, 'success').then(showManageMapelModal);
            });
        }

        function deleteMapel(encodedName) {
            const name = decodeURIComponent(encodedName);
            const usageCount = db.grades.filter(g => g.Subject === name).length
                + db.journal.filter(j => j.Subject === name).length
                + db.schedule.filter(s => s.Subject === name).length;
            const warningText = usageCount > 0
                ? `Mata pelajaran ini masih dipakai di ${usageCount} data Nilai/Jurnal/Jadwal. Data tersebut TIDAK akan terhapus, hanya mata pelajarannya yang tidak akan muncul lagi di pilihan dropdown.`
                : 'Mata pelajaran ini akan dihapus dari daftar.';

            Swal.fire({
                title: `Hapus "${escapeHtml(name)}"?`, text: warningText, icon: 'warning',
                showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    db.subjects = db.subjects.filter(m => m !== name);
                    populateMapelSelects(); syncDB();
                    Swal.fire('Terhapus!', `Mata Pelajaran "${name}" berhasil dihapus.`, 'success').then(showManageMapelModal);
                }
            });
        }

        // --- Dashboard / Beranda ---
        let dashboardTimer; 
        function renderBeranda() {
            document.getElementById('stat-siswa').innerText = db.students.length;
            document.getElementById('stat-kelas').innerText = db.classes.length;
            let tH = 0, tS = 0, tI = 0, tA = 0;
            db.attendance.forEach(rec => {
                try { let attData = typeof rec.AttendanceData === 'string' ? JSON.parse(rec.AttendanceData) : rec.AttendanceData;
                    for (let key in attData) { attData[key].forEach(val => { if(val === 'H') tH++; else if(val === 'S') tS++; else if(val === 'I') tI++; else if(val === 'A') tA++; }); }
                } catch(e) {}
            });
            let totalAtt = tH + tS + tI + tA;
            document.getElementById('stat-hadir').innerText = (totalAtt > 0 ? Math.round((tH / totalAtt) * 100) : 0) + "%";
            
            function updateLiveSchedule() {
                const daysIndo = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'], todayStr = daysIndo[new Date().getDay()];
                let todaySchedules = db.schedule.filter(s => s.Day === todayStr).sort((a, b) => a.TimeStart.localeCompare(b.TimeStart));
                const dashboardJadwal = document.getElementById('dashboard-jadwal');
                if (todaySchedules.length === 0) { dashboardJadwal.innerHTML = `<li class="list-group-item text-center text-muted py-3">Tidak ada jadwal hari ini (${todayStr})</li>`; } 
                else {
                    let htmlJadwal = ''; const now = new Date(), currentTime = now.getHours().toString().padStart(2, '0') + ':' + now.getMinutes().toString().padStart(2, '0');
                    todaySchedules.forEach(s => {
                        let className = s.ClassID === 'All' ? 'Semua Kelas' : (db.classes.find(c => c.ClassID === s.ClassID)?.ClassName || s.ClassID);
                        let statusBadge = currentTime > s.TimeEnd ? `<span class="badge bg-secondary rounded-pill">Selesai</span>` : (currentTime >= s.TimeStart && currentTime <= s.TimeEnd ? `<span class="badge bg-success rounded-pill placeholder-glow"><span class="placeholder col-12 px-2">Berlangsung</span></span>` : `<span class="badge bg-primary rounded-pill">Akan Datang</span>`);
                        htmlJadwal += `<li class="list-group-item d-flex justify-content-between align-items-center py-3"><div><span class="text-warning fw-bold me-2 fs-5"><i class="bi bi-clock"></i> ${s.TimeStart}</span> <span class="fw-bold">${s.Subject}</span> <small class="text-muted ms-1">(${className})</small></div>${statusBadge}</li>`;
                    });
                    dashboardJadwal.innerHTML = htmlJadwal;
                }
            }
            updateLiveSchedule();
            if(dashboardTimer) clearInterval(dashboardTimer);
            dashboardTimer = setInterval(updateLiveSchedule, 60000);

            if(chartHadir) chartHadir.destroy();
            chartHadir = new Chart(document.getElementById('chartKehadiran').getContext('2d'), {
                type: 'doughnut', data: { labels: ['Hadir', 'Sakit', 'Izin', 'Alpha'], datasets: [{ data: totalAtt > 0 ? [tH, tS, tI, tA] : [1, 0, 0, 0], backgroundColor: ['#198754', '#0dcaf0', '#ffc107', '#dc3545'] }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right' } } }
            });
        }

        // --- Class & Student Management ---
        function showAddClassModal() {
            Swal.fire({ title: 'Tambah Kelas Baru', input: 'text', inputPlaceholder: 'Nama Kelas (cth: Kelas 10-C)', showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal' }).then((result) => {
                if (result.isConfirmed && result.value) {
                    db.classes.push({ClassID: "C" + (db.classes.length + 1), ClassName: result.value}); populateClassSelects(); syncDB();
                    Swal.fire('Tersimpan!', `Kelas ${result.value} berhasil ditambahkan.`, 'success');
                }
            });
        }

        // Fungsi baru untuk menghapus Kelas beserta isinya
        function deleteClass() {
            const cid = document.getElementById('select-kelas-manajemen').value;
            if (!cid || cid === 'All') return Swal.fire('Peringatan', 'Pilih kelas yang valid untuk dihapus.', 'warning');
            
            const className = db.classes.find(c => c.ClassID === cid)?.ClassName || cid;

            Swal.fire({
                title: `Hapus Kelas ${className}?`,
                text: "Semua data siswa, absensi, dan nilai di kelas ini akan ikut terhapus secara permanen!",
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: '#d33',
                confirmButtonText: 'Ya, Hapus Kelas!'
            }).then((result) => {
                if (result.isConfirmed) {
                    // Hapus data kelas
                    db.classes = db.classes.filter(c => c.ClassID !== cid);
                    // Hapus semua siswa di kelas ini
                    db.students = db.students.filter(s => s.ClassID !== cid);

                    // Membersihkan data turunan secara permanen
                    db.attendance = db.attendance.filter(a => a.ClassID !== cid);
                    db.grades = db.grades.filter(g => g.ClassID !== cid);
                    db.journal = db.journal.filter(j => j.ClassID !== cid);
                    db.schedule = db.schedule.filter(s => s.ClassID !== cid);
                    db.koku = db.koku.filter(k => k.ClassID !== cid);
                    db.poin = db.poin.filter(p => p.ClassID !== cid);

                    syncDB(); 
                    populateClassSelects(); 
                    
                    // Reset tabel
                    document.getElementById('select-kelas-manajemen').value = '';
                    renderStudentsTable(); 
                    
                    Swal.fire('Terhapus!', `Kelas ${className} beserta data siswanya telah berhasil dihapus.`, 'success');
                }
            });
        }

        function showAddStudentModal() {
            const cid = document.getElementById('select-kelas-manajemen').value;
            if (!cid || cid === 'All') return Swal.fire('Peringatan', 'Pilih kelas terlebih dahulu.', 'warning');
            Swal.fire({ 
                title: 'Tambah Siswa Baru', 
                html: `<div class="text-start mb-2"><label class="form-label small fw-bold">Nama Siswa</label><input type="text" id="swal-input-name" class="form-control" placeholder="Nama Lengkap Siswa"></div>
                       <div class="text-start"><label class="form-label small fw-bold">No. WhatsApp Ortu</label><input type="text" id="swal-input-phone" class="form-control" placeholder="Cth: 0812..."></div>`,
                showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal',
                preConfirm: () => {
                    let name = document.getElementById('swal-input-name').value;
                    let phone = document.getElementById('swal-input-phone').value;
                    if(!name) Swal.showValidationMessage('Nama siswa wajib diisi');
                    return [name, phone];
                } 
            }).then((result) => {
                if (result.isConfirmed && result.value[0]) {
                    db.students.push({ StudentID: "S" + Date.now(), ClassID: cid, StudentName: result.value[0].trim(), ParentPhone: result.value[1].trim() }); 
                    syncDB(); renderStudentsTable(); Swal.fire('Tersimpan!', `Siswa berhasil ditambahkan.`, 'success');
                }
            });
        }

        function editStudent(studentId) {
            const student = db.students.find(s => s.StudentID === studentId);
            if (!student) return;

            Swal.fire({ 
                title: 'Edit Data Siswa', 
                html: `<div class="text-start mb-2"><label class="form-label small fw-bold">Nama Siswa</label><input type="text" id="swal-edit-name" class="form-control" placeholder="Nama Lengkap Siswa" value="${student.StudentName}"></div>
                       <div class="text-start"><label class="form-label small fw-bold">No. WhatsApp</label><input type="text" id="swal-edit-phone" class="form-control" placeholder="Cth: 0812..." value="${student.ParentPhone || ''}"></div>`,
                showCancelButton: true, confirmButtonText: 'Simpan', cancelButtonText: 'Batal',
                preConfirm: () => {
                    let name = document.getElementById('swal-edit-name').value;
                    let phone = document.getElementById('swal-edit-phone').value;
                    if(!name) Swal.showValidationMessage('Nama siswa wajib diisi');
                    return [name, phone];
                } 
            }).then((result) => {
                if (result.isConfirmed && result.value[0]) {
                    const idx = db.students.findIndex(s => s.StudentID === studentId);
                    if (idx > -1) {
                        db.students[idx].StudentName = result.value[0].trim();
                        db.students[idx].ParentPhone = result.value[1].trim();
                        syncDB(); 
                        renderStudentsTable(); 
                        
                        // Render ulang tabel lain jika sedang terbuka
                        if(document.getElementById('view-absensi').classList.contains('active')) loadAbsensiData();
                        if(document.getElementById('view-nilai').classList.contains('active')) loadNilaiData();

                        Swal.fire('Tersimpan!', `Data siswa berhasil diperbarui.`, 'success');
                    }
                }
            });
        }

        function showUploadExcelModal() {
            const cid = document.getElementById('select-kelas-manajemen').value;
            if (!cid || cid === 'All') return Swal.fire('Peringatan', 'Pilih kelas terlebih dahulu.', 'warning');
            const className = db.classes.find(c => c.ClassID === cid)?.ClassName || cid;
            Swal.fire({
                title: `Upload Siswa ke ${className}`, html: `<div class="text-start"><p class="text-muted small mb-3">Kolom wajib: <b>Nama Siswa</b>.<br>Kolom opsional: <b>No HP</b> (untuk WA ortu).</p><input type="file" id="excel-file-upload" class="form-control" accept=".xlsx, .xls"></div>`,
                showCancelButton: true, confirmButtonText: 'Proses Upload', cancelButtonText: 'Batal',
                preConfirm: () => { const file = document.getElementById('excel-file-upload').files[0]; if (!file) Swal.showValidationMessage('Pilih file Excel'); return file; }
            }).then((result) => { if (result.isConfirmed && result.value) processExcelUpload(result.value, cid); });
        }

        async function processExcelUpload(file, classId) {
            await ensureXLSX();
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result), workbook = XLSX.read(data, {type: 'array'}), json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                    if (json.length === 0) return Swal.fire('Error', 'File kosong.', 'error');
                    const nameKey = Object.keys(json[0]).find(key => key.toLowerCase().includes('nama'));
                    if (!nameKey) return Swal.fire('Error', 'Kolom "Nama" tidak ditemukan di Excel.', 'error');
                    
                    const phoneKey = Object.keys(json[0]).find(key => {
                        let k = key.toLowerCase();
                        return k.includes('hp') || k.includes('wa') || k.includes('telepon') || k.includes('whatsapp') || k.includes('phone');
                    });

                    let countAdded = 0;
                    json.forEach(row => { 
                        if (row[nameKey]) { 
                            let parentPhone = phoneKey && row[phoneKey] ? row[phoneKey].toString().trim() : "";
                            db.students.push({ 
                                StudentID: "S" + Date.now() + countAdded, 
                                ClassID: classId, 
                                StudentName: row[nameKey].toString().trim(),
                                ParentPhone: parentPhone
                            }); 
                            countAdded++; 
                        } 
                    });
                    syncDB(); renderStudentsTable(); Swal.fire('Berhasil!', `${countAdded} data siswa diimpor.`, 'success');
                } catch (error) { Swal.fire('Error', 'Format Excel salah.', 'error'); }
            }; reader.readAsArrayBuffer(file);
        }

        function renderStudentsTable() {
            const cid = document.getElementById('select-kelas-manajemen').value, tbody = document.getElementById('table-students-body');
            const btnDeleteClass = document.getElementById('btn-hapus-kelas');
            
            // Tampilkan atau sembunyikan tombol Hapus Kelas
            if(!cid || cid === 'All') {
                if(btnDeleteClass) btnDeleteClass.style.display = 'none';
                return tbody.innerHTML = '<tr><td colspan="4" class="text-center">Pilih kelas</td></tr>';
            }
            if(btnDeleteClass) btnDeleteClass.style.display = 'inline-block';
            
            const stds = db.students.filter(s => String(s.ClassID) === String(cid));
            if(stds.length === 0) return tbody.innerHTML = '<tr><td colspan="4" class="text-center">Belum ada siswa</td></tr>';
            
            let html = '';
            stds.forEach((s, i) => { 
                let phoneText = s.ParentPhone ? `<br><small class="text-muted"><i class="bi bi-whatsapp text-success"></i> ${s.ParentPhone}</small>` : '<br><small class="text-danger" style="font-size:0.7rem;"><i>Belum ada No. HP</i></small>';
                let chkDisabled = s.ParentPhone ? '' : 'disabled';
                html += `<tr>
                    <td class="text-center align-middle"><input type="checkbox" class="chk-std form-check-input" value="${s.StudentID}" ${chkDisabled}></td>
                    <td class="text-center align-middle">${i+1}</td>
                    <td class="align-middle">${s.StudentName}${phoneText}</td>
                    <td class="text-center align-middle">
                        <button class="btn btn-sm btn-primary me-1 mb-1" onclick="editStudent('${s.StudentID}')" title="Edit Siswa"><i class="bi bi-pencil-square"></i></button>
                        <button class="btn btn-sm btn-danger mb-1" onclick="deleteStudent('${s.StudentID}')" title="Hapus Siswa"><i class="bi bi-trash"></i></button>
                    </td>
                </tr>`; 
            });
            tbody.innerHTML = html;
        }

        function deleteStudent(studentId) {
            Swal.fire({ title: 'Hapus Siswa?', text: "Yakin ingin menghapus?", icon: 'warning', showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Ya, Hapus!' }).then((result) => {
                if (result.isConfirmed) { db.students = db.students.filter(s => s.StudentID !== studentId); syncDB(); renderStudentsTable(); Swal.fire('Terhapus!', 'Data dihapus.', 'success'); }
            });
        }

        // --- Absensi ---
        function loadAbsensiData() {
            const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
            if(!cid) { document.getElementById('absensi-chart-container').style.display = 'none'; return; }
            
            const stds = db.students.filter(s => String(s.ClassID) === String(cid));
            const record = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month));
            
            let savedDates = [], savedAttendance = {};
            if (record) { try { savedDates = JSON.parse(record.DatesData); savedAttendance = JSON.parse(record.AttendanceData); } catch(e) {} }
            
            if (savedDates.length === 0) savedDates = [""];
            let colCount = savedDates.length;

            let thead = '<th>Nama Siswa</th>';
            for(let i=0; i<colCount; i++) {
                thead += `<th><input type="date" class="date-input-header date-absen-col" data-idx="${i}" value="${savedDates[i] || ""}" onchange="saveAbsensi(true)">
                    ${colCount > 1 ? `<button type="button" class="btn btn-sm btn-outline-danger mt-1 py-0 px-1" onclick="deleteAbsenColumn(${i})" title="Hapus Kolom Ini"><i class="bi bi-trash"></i></button>` : ''}
                </th>`;
            }
            document.getElementById('absen-thead').innerHTML = thead;

            let tbody = '';
            stds.forEach(s => {
                tbody += `<tr><td class="fw-bold d-flex align-items-center justify-content-between">${s.StudentName}
                <button class="btn btn-sm btn-outline-success py-0 px-1" onclick="sendWAAbsen('${s.StudentID}')" title="Kirim WA Orang Tua"><i class="bi bi-whatsapp"></i></button>
                </td>`;
                let sAtt = savedAttendance[s.StudentID] || Array(colCount).fill('-');
                for(let i=0; i<colCount; i++) {
                    let val = sAtt[i] || '-', color = val !== '-' ? `status-${val}` : '';
                    tbody += `<td><select class="status-select select-absen-val ${color}" data-sid="${s.StudentID}" data-idx="${i}" onchange="updateSelectColor(this)">
                        <option value="H" ${val === 'H' ? 'selected' : ''}>H</option><option value="S" ${val === 'S' ? 'selected' : ''}>S</option>
                        <option value="I" ${val === 'I' ? 'selected' : ''}>I</option><option value="A" ${val === 'A' ? 'selected' : ''}>A</option><option value="-" ${val === '-' ? 'selected' : ''}>-</option>
                    </select></td>`;
                }
                tbody += `</tr>`;
            });
            document.getElementById('absen-tbody').innerHTML = tbody; document.getElementById('absensi-chart-container').style.display = 'block'; renderChartAbsensi();
        }

        function addAbsenColumn() {
            const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
            if(!cid) return Swal.fire('Error', 'Pilih kelas terlebih dahulu.', 'error');
            saveAbsensi(true);
            let record = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month));
            if (record) {
                let dates = JSON.parse(record.DatesData); dates.push(""); record.DatesData = JSON.stringify(dates);
                let att = JSON.parse(record.AttendanceData); for(let sid in att) { att[sid].push("-"); } record.AttendanceData = JSON.stringify(att); syncDB();
            }
            loadAbsensiData();
        }

        function deleteAbsenColumn(idx) {
            const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
            if(!cid) return;
            saveAbsensi(true); // pastikan perubahan yang belum tersimpan tidak ikut hilang
            let record = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month));
            if (!record) return;
            let dates = []; try { dates = JSON.parse(record.DatesData); } catch(e) {}
            if (dates.length <= 1) return Swal.fire('Tidak Bisa', 'Minimal harus ada 1 kolom tanggal.', 'warning');

            Swal.fire({
                title: 'Hapus Kolom Ini?', text: 'Seluruh data absensi pada kolom tanggal ini akan dihapus.', icon: 'warning',
                showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
            }).then((result) => {
                if (!result.isConfirmed) return;
                dates.splice(idx, 1);
                record.DatesData = JSON.stringify(dates);
                let att = {}; try { att = JSON.parse(record.AttendanceData); } catch(e) {}
                for (let sid in att) { att[sid].splice(idx, 1); }
                record.AttendanceData = JSON.stringify(att);
                syncDB(); loadAbsensiData();
                Swal.fire('Terhapus!', 'Kolom berhasil dihapus.', 'success');
            });
        }

        function updateSelectColor(sel) { sel.className = `status-select select-absen-val status-${sel.value}`; renderChartAbsensi(); saveAbsensi(true); }

        function renderChartAbsensi() {
            let tH = 0, tS = 0, tI = 0, tA = 0;
            document.querySelectorAll('.select-absen-val').forEach(sel => { if (sel.value === 'H') tH++; else if (sel.value === 'S') tS++; else if (sel.value === 'I') tI++; else if (sel.value === 'A') tA++; });
            if (chartAbsensiObj) chartAbsensiObj.destroy();
            chartAbsensiObj = new Chart(document.getElementById('chartAbsensi').getContext('2d'), {
                type: 'bar', data: { labels: ['Hadir', 'Sakit', 'Izin', 'Alpha'], datasets: [{ label: 'Kehadiran', data: [tH, tS, tI, tA], backgroundColor: ['#198754', '#0dcaf0', '#ffc107', '#dc3545'], borderRadius: 5 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
            });
        }

        function saveAbsensi(isSilent = false) {
            const cid = document.getElementById('absen-kelas').value, month = document.getElementById('absen-bulan').value;
            if(!cid) { if (!isSilent) Swal.fire('Error', 'Pilih kelas', 'error'); return; }
            let customDates = []; document.querySelectorAll('.date-absen-col').forEach(i => customDates.push(i.value));
            let attData = {};
            db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => { let st = []; document.querySelectorAll(`.select-absen-val[data-sid="${s.StudentID}"]`).forEach(sel => st.push(sel.value)); attData[s.StudentID] = st; });

            let idx = db.attendance.findIndex(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(month));
            if (idx >= 0) { db.attendance[idx].DatesData = JSON.stringify(customDates); db.attendance[idx].AttendanceData = JSON.stringify(attData); }
            else db.attendance.push({ RecordID: "A" + Date.now(), ClassID: cid, MonthYear: month, DatesData: JSON.stringify(customDates), AttendanceData: JSON.stringify(attData) });
            syncDB(); if (!isSilent) Swal.fire({icon: 'success', title: 'Tersimpan', text: 'Data Absensi disimpan', timer: 1500, showConfirmButton: false});
        }

        async function downloadAbsenExcel(semester) {
            const cid = document.getElementById('absen-kelas').value;
            if (!cid) return Swal.fire('Error', 'Pilih kelas', 'error');
            saveAbsensi(true); const className = db.classes.find(c => c.ClassID === cid)?.ClassName || cid;
            Swal.fire({icon: 'info', title: 'Mengekspor', text: `Menyiapkan Excel...`, timer: 1500});
            await ensureXLSX();
            const wb = XLSX.utils.book_new(), stds = db.students.filter(s => String(s.ClassID) === String(cid));
            const targetMonths = semester === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];
            const monthNames = {'01':'Jan', '02':'Feb', '03':'Mar', '04':'Apr', '05':'Mei', '06':'Jun', '07':'Jul', '08':'Agu', '09':'Sep', '10':'Okt', '11':'Nov', '12':'Des'};
            let rekapSemester = {}; stds.forEach(s => rekapSemester[s.StudentID] = { Name: s.StudentName, H: 0, S: 0, I: 0, A: 0 });

            targetMonths.forEach(m => {
                const record = db.attendance.find(a => String(a.ClassID) === String(cid) && String(a.MonthYear) === String(m));
                let sDates = [], sAtt = {};
                if (record) { try { sDates = JSON.parse(record.DatesData); sAtt = JSON.parse(record.AttendanceData); } catch(e){} }
                let maxCols = sDates.length > 0 ? sDates.length : 1;
                let headDates = Array.from({length: maxCols}, (_, i) => sDates[i] || `K-${i+1}`);
                let ws_data = [ ["Nama Siswa", ...headDates, "Total H", "Total S", "Total I", "Total A"] ];

                stds.forEach(s => {
                    let row = [s.StudentName], att = sAtt[s.StudentID] || Array(maxCols).fill('-');
                    let mH=0, mS=0, mI=0, mA=0; 
                    for(let i=0; i<maxCols; i++) { 
                        let val = att[i] || '-';
                        row.push(val); 
                        if (val==='H'){mH++;rekapSemester[s.StudentID].H++;} 
                        if (val==='S'){mS++;rekapSemester[s.StudentID].S++;} 
                        if (val==='I'){mI++;rekapSemester[s.StudentID].I++;} 
                        if (val==='A'){mA++;rekapSemester[s.StudentID].A++;} 
                    }
                    row.push(mH, mS, mI, mA); ws_data.push(row);
                });
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ws_data), monthNames[m]);
            });

            let rekap_data = [ ["Nama Siswa", "Total Hadir (H)", "Total Sakit (S)", "Total Izin (I)", "Total Alpha (A)"] ];
            stds.forEach(s => { let r = rekapSemester[s.StudentID]; rekap_data.push([r.Name, r.H, r.S, r.I, r.A]); });
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rekap_data), "Rekap"); XLSX.writeFile(wb, `Absensi_${className}_${semester}.xlsx`);
        }

        // --- Data Nilai ---
        function loadNilaiData() {
            const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value, jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value, container = document.getElementById('container-table-nilai');
            if(!cid || !mapel) { container.innerHTML = '<div class="alert alert-secondary text-center">Silakan pilih Kelas dan Mata Pelajaran.</div>'; document.getElementById('nilai-stats-container').style.display = 'none'; document.getElementById('btn-add-nilai-col').style.display = 'none'; return; }
            
            const stds = db.students.filter(s => String(s.ClassID) === String(cid));
            const record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === String(jenis) && (jenis !== 'Formatif' || String(g.MonthYear) === String(bulan)));
            
            let sDates = [], sGrades = {}, sCakupan = [];
            if(record) { try { sDates = JSON.parse(record.DatesData) || []; sGrades = JSON.parse(record.GradesData) || {}; sCakupan = JSON.parse(record.Cakupan) || []; } catch(e) {} }
            
            let html = '';
            if(jenis === 'Formatif') {
                document.getElementById('btn-add-nilai-col').style.display = 'inline-block';
                if (sDates.length === 0) { sDates = [""]; sCakupan = [""]; }
                let colCount = sDates.length;

                html += '<table class="table table-bordered table-dynamic" id="table-nilai-input"><thead><tr class="table-success"><th style="min-width:210px; vertical-align:middle;">Nama Siswa</th>';
                for(let i=0; i<colCount; i++) {
                    html += `<th><input type="date" class="date-input-header date-nilai-col mb-1" value="${sDates[i] || ""}" onchange="saveNilai(true)"><input type="text" class="form-control form-control-sm cakupan-nilai-col" style="font-size:0.75rem; width:125px; margin:0 auto;" placeholder="Materi" value="${sCakupan[i] || ""}" onchange="saveNilai(true)">
                        ${colCount > 1 ? `<button type="button" class="btn btn-sm btn-outline-danger mt-1 py-0 px-1" onclick="deleteNilaiColumn(${i})" title="Hapus Kolom Ini"><i class="bi bi-trash"></i></button>` : ''}
                    </th>`;
                }
                html += '</tr></thead><tbody>';
                stds.forEach(s => {
                    html += `<tr><td class="fw-bold d-flex align-items-center justify-content-between">${s.StudentName}
                    <button class="btn btn-sm btn-outline-success py-0 px-1" onclick="sendWANilai('${s.StudentID}')" title="Kirim WA Orang Tua"><i class="bi bi-whatsapp"></i></button>
                    </td>`;
                    let grades = sGrades[s.StudentID] || Array(colCount).fill('');
                    for(let i=0; i<colCount; i++) html += `<td><input type="number" class="form-control form-control-sm px-1 text-center input-nilai-val" style="width:60px; margin:0 auto;" data-sid="${s.StudentID}" value="${grades[i] || ''}" oninput="calculateStatsNilai()"></td>`;
                    html += '</tr>';
                });
                html += '</tbody></table>'; document.getElementById('nilai-bulan').disabled = false;
            } else {
                document.getElementById('btn-add-nilai-col').style.display = 'none';
                html += `<table class="table table-bordered" id="table-nilai-input"><thead class="table-primary"><tr><th width="30%">Nama Siswa</th><th>Nilai ${jenis}</th></tr></thead><tbody>`;
                stds.forEach(s => { 
                    html += `<tr><td class="fw-bold d-flex align-items-center justify-content-between">${s.StudentName}
                    <button class="btn btn-sm btn-outline-success py-0 px-1" onclick="sendWANilai('${s.StudentID}')" title="Kirim WA Orang Tua"><i class="bi bi-whatsapp"></i></button>
                    </td><td><input type="number" class="form-control w-50 input-nilai-val" data-sid="${s.StudentID}" value="${sGrades[s.StudentID] ? sGrades[s.StudentID][0] : ''}" oninput="calculateStatsNilai()"></td></tr>`; 
                });
                html += '</tbody></table>'; document.getElementById('nilai-bulan').disabled = true;
            }
            container.innerHTML = html; document.getElementById('nilai-stats-container').style.display = 'flex'; document.getElementById('nilai-chart-container').style.display = 'block'; calculateStatsNilai(true);
        }

        function addNilaiColumn() {
            const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value, jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value;
            if(!cid || !mapel) return Swal.fire('Error', 'Pilih kelas & mapel', 'error');
            if(jenis !== 'Formatif') return;
            saveNilai(true);
            let record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === String(jenis) && String(g.MonthYear) === String(bulan));
            if (record) {
                let dates = JSON.parse(record.DatesData); dates.push(""); record.DatesData = JSON.stringify(dates);
                let caku = JSON.parse(record.Cakupan); caku.push(""); record.Cakupan = JSON.stringify(caku);
                let gd = JSON.parse(record.GradesData); for(let sid in gd) { gd[sid].push(""); } record.GradesData = JSON.stringify(gd); syncDB();
            }
            loadNilaiData();
        }

        function deleteNilaiColumn(idx) {
            const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value, jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value;
            if(!cid || !mapel || jenis !== 'Formatif') return;
            saveNilai(true); // pastikan perubahan yang belum tersimpan tidak ikut hilang
            let record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === String(jenis) && String(g.MonthYear) === String(bulan));
            if (!record) return;
            let dates = []; try { dates = JSON.parse(record.DatesData); } catch(e) {}
            if (dates.length <= 1) return Swal.fire('Tidak Bisa', 'Minimal harus ada 1 kolom penilaian.', 'warning');

            Swal.fire({
                title: 'Hapus Kolom Ini?', text: 'Seluruh nilai siswa pada kolom ini akan dihapus.', icon: 'warning',
                showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
            }).then((result) => {
                if (!result.isConfirmed) return;
                dates.splice(idx, 1);
                record.DatesData = JSON.stringify(dates);
                let caku = []; try { caku = JSON.parse(record.Cakupan); } catch(e) {}
                caku.splice(idx, 1);
                record.Cakupan = JSON.stringify(caku);
                let gd = {}; try { gd = JSON.parse(record.GradesData); } catch(e) {}
                for (let sid in gd) { gd[sid].splice(idx, 1); }
                record.GradesData = JSON.stringify(gd);
                syncDB(); loadNilaiData();
                Swal.fire('Terhapus!', 'Kolom berhasil dihapus.', 'success');
            });
        }

        function calculateStatsNilai(isInit = false) {
            const cid = document.getElementById('nilai-kelas').value, jenis = document.getElementById('nilai-jenis').value, kkmValue = parseFloat(document.getElementById('input-kkm').value) || 75; 
            let stdAvgs = [], names = [], scores = [], kkmNames = [];
            db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => {
                let sum = 0, count = 0; document.querySelectorAll(`.input-nilai-val[data-sid="${s.StudentID}"]`).forEach(inp => { let val = parseFloat(inp.value); if(!isNaN(val)) { sum += val; count++; } });
                let avg = count > 0 ? (sum / count) : 0; stdAvgs.push({name: s.StudentName, avg: avg}); names.push(s.StudentName.split(' ')[0]); scores.push(avg.toFixed(1));
                if (count > 0 && avg < kkmValue) kkmNames.push(s.StudentName);
            });
            let highest = stdAvgs.reduce((max, obj) => (obj.avg > max.avg ? obj : max), {name: '-', avg: 0});
            document.getElementById('stat-nilai-max-name').innerText = highest.avg > 0 ? highest.name : '-'; document.getElementById('stat-nilai-max-val').innerText = highest.avg > 0 ? `Nilai: ${highest.avg.toFixed(1)}` : 'Nilai: 0';
            document.getElementById('stat-nilai-kkm-count').innerText = `${kkmNames.length} Siswa`; document.getElementById('stat-nilai-kkm-names').innerText = kkmNames.length > 0 ? kkmNames.join(', ') : 'Semua tuntas!';
            if(chartPrestasi) chartPrestasi.destroy();
            chartPrestasi = new Chart(document.getElementById('chartNilai').getContext('2d'), { type: 'bar', data: { labels: names, datasets: [{ label: `Rata-rata ${jenis}`, data: scores, backgroundColor: 'rgba(0, 123, 255, 0.7)' }] }, options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true, max: 100 } } } });
            if(!isInit) saveNilai(true); 
        }

        function saveNilai(isSilent = false) {
            const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value, jenis = document.getElementById('nilai-jenis').value, bulan = document.getElementById('nilai-bulan').value;
            if(!cid || !mapel) { if(!isSilent) Swal.fire('Error', 'Pilih Kelas & Mapel', 'error'); return; }
            let customDates = [], customCakupan = [], gradesData = {};
            if(jenis === 'Formatif') {
                document.querySelectorAll('.date-nilai-col').forEach(input => customDates.push(input.value)); document.querySelectorAll('.cakupan-nilai-col').forEach(input => customCakupan.push(input.value));
                db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => { let gArray = []; document.querySelectorAll(`.input-nilai-val[data-sid="${s.StudentID}"]`).forEach(inp => gArray.push(inp.value)); gradesData[s.StudentID] = gArray; });
            } else { db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => { let inp = document.querySelector(`.input-nilai-val[data-sid="${s.StudentID}"]`); gradesData[s.StudentID] = [inp ? inp.value : '']; }); }

            let idx = db.grades.findIndex(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === String(jenis) && (jenis !== 'Formatif' || String(g.MonthYear) === String(bulan)));
            if (idx >= 0) { db.grades[idx].Cakupan = JSON.stringify(customCakupan); db.grades[idx].DatesData = JSON.stringify(customDates); db.grades[idx].GradesData = JSON.stringify(gradesData); }
            else db.grades.push({ RecordID: "G" + Date.now(), ClassID: cid, Subject: mapel, Cakupan: JSON.stringify(customCakupan), Type: jenis, MonthYear: jenis === 'Formatif' ? bulan : null, DatesData: JSON.stringify(customDates), GradesData: JSON.stringify(gradesData) });
            syncDB(); if (!isSilent) Swal.fire({icon: 'success', title: 'Tersimpan', text: 'Data disimpan!', timer: 1500, showConfirmButton: false});
        }

        async function downloadNilaiExcel(semester) {
            saveNilai(true); const cid = document.getElementById('nilai-kelas').value, mapel = document.getElementById('nilai-mapel').value;
            if (!cid || !mapel) return Swal.fire('Error', 'Pilih kelas & mapel', 'error');
            const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
            Swal.fire({icon: 'info', title: 'Mengekspor', text: `Menyiapkan Excel...`, timer: 1500});
            await ensureXLSX();
            const wb = XLSX.utils.book_new(), stds = db.students.filter(s => String(s.ClassID) === String(cid));
            const targetMonths = semester === 'ganjil' ? ['07', '08', '09', '10', '11', '12'] : ['01', '02', '03', '04', '05', '06'];
            const monthNames = {'01':'Jan', '02':'Feb', '03':'Mar', '04':'Apr', '05':'Mei', '06':'Jun', '07':'Jul', '08':'Agu', '09':'Sep', '10':'Okt', '11':'Nov', '12':'Des'};
            let rekapSemester = {}; stds.forEach(s => rekapSemester[s.StudentID] = { Name: s.StudentName, SumFormatif: 0, CountFormatif: 0, SLM: 0, SAS: 0 });

            targetMonths.forEach(m => {
                const record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'Formatif' && String(g.MonthYear) === String(m));
                let sDates = [], sGrades = {}, sCakupan = [];
                if (record) { try { sDates = JSON.parse(record.DatesData)||[]; sGrades = JSON.parse(record.GradesData)||{}; sCakupan = JSON.parse(record.Cakupan)||[]; } catch(e) {} }
                let maxCols = sDates.length > 0 ? sDates.length : 1;
                let headDates = Array.from({length: maxCols}, (_, i) => sDates[i] || `N-${i+1}`);
                let headCakupan = Array.from({length: maxCols}, (_, i) => sCakupan[i] || '-');
                let ws_data = [ [`Mapel: ${mapel}`], [], ["Tanggal", ...headDates, ""], ["Cakupan", ...headCakupan, "Rata-rata"] ];
                stds.forEach(s => {
                    let row = [s.StudentName], grades = sGrades[s.StudentID] || Array(maxCols).fill(''), sum = 0, count = 0;
                    for(let i=0; i<maxCols; i++) {
                        let gradeVal = grades[i] || '';
                        row.push(gradeVal);
                        if(gradeVal !== "" && !isNaN(parseFloat(gradeVal))) { let num = parseFloat(gradeVal); sum += num; count++; rekapSemester[s.StudentID].SumFormatif += num; rekapSemester[s.StudentID].CountFormatif++; }
                    }
                    row.push(count > 0 ? (sum/count).toFixed(1) : 0); ws_data.push(row);
                });
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ws_data), monthNames[m]);
            });

            const recSLM = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifLM'); let slmData = {}; if(recSLM) try { slmData = JSON.parse(recSLM.GradesData); } catch(e){}
            const recSAS = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifAS'); let sasData = {}; if(recSAS) try { sasData = JSON.parse(recSAS.GradesData); } catch(e){}
            let rekap_data = [ [`Rapor - ${mapel}`], [], ["Nama", "Rata Formatif", "SLM", "SAS", "Rapor"] ];
            stds.forEach(s => {
                let r = rekapSemester[s.StudentID], avg = r.CountFormatif > 0 ? (r.SumFormatif / r.CountFormatif) : 0;
                let slm = parseFloat(slmData[s.StudentID]?.[0] || 0), sas = parseFloat(sasData[s.StudentID]?.[0] || 0);
                rekap_data.push([ r.Name, avg.toFixed(1), slm.toFixed(1), sas.toFixed(1), ((avg+slm+sas)/3).toFixed(1) ]);
            });
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rekap_data), "Rekap"); XLSX.writeFile(wb, `Nilai_${mapel}_${semester}.xlsx`);
        }

        async function downloadNilaiPDF() {
            saveNilai(true); const cid = document.getElementById('nilai-kelas').value; const mapel = document.getElementById('nilai-mapel').value;
            if (!cid || !mapel) return Swal.fire('Error', 'Pilih kelas & mapel', 'error');
            await ensurePDF();
            const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
            let rekapSemester = {}; db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => rekapSemester[s.StudentID] = { Name: s.StudentName, SumFormatif: 0, CountFormatif: 0 });

            ['01','02','03','04','05','06','07','08','09','10','11','12'].forEach(m => {
                const record = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'Formatif' && String(g.MonthYear) === String(m));
                if (record) {
                    let sG = {}; try { sG = JSON.parse(record.GradesData); } catch(e) {}
                    db.students.filter(s => String(s.ClassID) === String(cid)).forEach(s => { 
                        let grades = sG[s.StudentID] || []; 
                        for(let i=0; i<grades.length; i++) {
                            if(grades[i] !== "" && !isNaN(parseFloat(grades[i]))) { 
                                rekapSemester[s.StudentID].SumFormatif += parseFloat(grades[i]); rekapSemester[s.StudentID].CountFormatif++; 
                            } 
                        }
                    });
                }
            });

            const recSLM = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifLM'); let slmD = {}; if(recSLM) try { slmD = JSON.parse(recSLM.GradesData); } catch(e){}
            const recSAS = db.grades.find(g => String(g.ClassID) === String(cid) && String(g.Subject) === String(mapel) && String(g.Type) === 'SumatifAS'); let sasD = {}; if(recSAS) try { sasD = JSON.parse(recSAS.GradesData); } catch(e){}

            const { jsPDF } = window.jspdf; const doc = new jsPDF();
            doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.text(`Rekapitulasi Nilai - ${mapel}`, 14, 16); doc.setFontSize(11); doc.setFont("helvetica", "normal"); doc.text(`Kelas: ${className}`, 14, 22);
            const tableData = db.students.filter(s => String(s.ClassID) === String(cid)).map(s => {
                let r = rekapSemester[s.StudentID], avg = r.CountFormatif > 0 ? (r.SumFormatif / r.CountFormatif) : 0;
                let slm = parseFloat(slmD[s.StudentID]?.[0] || 0), sas = parseFloat(sasD[s.StudentID]?.[0] || 0);
                return [ s.StudentName, avg.toFixed(1), slm.toFixed(1), sas.toFixed(1), ((avg+slm+sas)/3).toFixed(1) ];
            });
            doc.autoTable({ head: [['Nama Siswa', 'Rata Formatif', 'Sumatif LM', 'Sumatif AS', 'Nilai Rapor']], body: tableData, startY: 28, headStyles: { fillColor: [0, 123, 255] } });
            doc.save(`Rekap_${mapel}.pdf`);
        }

        // --- Kokurikuler ---
        function renderKokuDimensi() {
            let html = ''; ['Keimanan & Ketakwaan', 'Kewargaan', 'Penalaran Kritis', 'Kreativitas', 'Kolaborasi', 'Kemandirian', 'Kesehatan', 'Komunikasi'].forEach((d, i) => { html += `<div class="col-md-3 col-6"><div class="form-check"><input class="form-check-input dim-checkbox" type="checkbox" value="${d}" id="dim-chk-${i}"><label class="form-check-label" style="font-size:0.85rem" for="dim-chk-${i}">${d}</label></div></div>`; });
            document.getElementById('koku-dimensi-container').innerHTML = html; document.getElementById('koku-tanggal').value = new Date().toISOString().split('T')[0];
        }

        // Menampilkan daftar penilaian kokurikuler yang SUDAH TERSIMPAN untuk kelas yang dipilih,
        // supaya data yang sudah diinput sebelumnya tidak hilang dari tampilan (bisa dilihat/diedit lagi).
        function loadKokuHistory() {
            const cid = document.getElementById('koku-kelas').value;
            const section = document.getElementById('koku-history-section'), tbody = document.getElementById('koku-history-tbody');
            if (!cid) { section.style.display = 'none'; return; }

            const records = db.koku.filter(k => String(k.ClassID) === String(cid));
            if (records.length === 0) { section.style.display = 'none'; return; }

            section.style.display = 'block';
            let html = '';
            let filterOpts = '<option value="">Semua Projek (Kelas Ini)</option>';
            records.forEach(k => {
                let jumlah = 0;
                try { jumlah = Object.keys(JSON.parse(k.KokuData || '{}')).length; } catch(e) {}
                html += `<tr><td>${k.Tanggal || '-'}</td><td>${k.Jenis === '7KAIH' ? '7KAIH' : 'Lintas Mata Pelajaran'}</td><td class="fw-bold">${escapeHtml(k.Projek || '-')}</td><td>${jumlah} siswa</td>
                <td class="text-nowrap"><button class="btn btn-sm btn-primary me-1 mb-1" onclick="editKokuRecord('${k.RecordID}')" title="Lihat/Edit"><i class="bi bi-eye"></i></button><button class="btn btn-sm btn-danger mb-1" onclick="deleteKokuRecord('${k.RecordID}')" title="Hapus"><i class="bi bi-trash"></i></button></td></tr>`;
                filterOpts += `<option value="${k.RecordID}">${escapeHtml(k.Projek || '-')} (${k.Tanggal || '-'})</option>`;
            });
            tbody.innerHTML = html;
            document.getElementById('koku-export-filter').innerHTML = filterOpts;
        }

        function generateKokuTable() {
            const cid = document.getElementById('koku-kelas').value, p = document.getElementById('koku-projek').value, desc = document.getElementById('koku-deskripsi').value;
            let selectedDims = []; document.querySelectorAll('.dim-checkbox:checked').forEach(chk => selectedDims.push(chk.value));
            if(!cid || !p || !desc || selectedDims.length === 0) return Swal.fire('Peringatan', 'Harap isi form & minimal 1 dimensi.', 'warning');
            const stds = db.students.filter(s => String(s.ClassID) === String(cid)); if(stds.length === 0) return Swal.fire('Info', 'Tidak ada siswa.', 'info');
            editingKokuId = null; // ini selalu membuat penilaian BARU, bukan mengedit yang sudah ada
            window.activeKokuProject = { p, desc, selectedDims }; let html = '';
            stds.forEach((s, i) => {
                html += `<tr data-sid="${s.StudentID}" class="koku-row"><td>${i+1}</td><td class="fw-bold">${s.StudentName}</td><td><select class="form-select form-select-sm koku-cat-select" onchange="updateKokuNarasi(this)"><option value="">Pilih</option><option value="Sangat Baik">Sangat Baik</option><option value="Baik">Baik</option><option value="Cukup">Cukup</option><option value="Kurang">Kurang</option></select></td><td><small class="koku-narasi-text text-muted">Pilih kategori...</small></td></tr>`;
            });
            document.getElementById('koku-tbody').innerHTML = html; document.getElementById('koku-result-section').style.display = 'block';
        }

        function updateKokuNarasi(selectElem) {
            const cat = selectElem.value, row = selectElem.closest('tr'), narasiContainer = row.querySelector('.koku-narasi-text'), sName = row.cells[1].innerText;
            if(!cat) { narasiContainer.innerHTML = 'Pilih kategori...'; return; }
            const proj = window.activeKokuProject;
            let dimStr = proj.selectedDims.length > 1 ? proj.selectedDims.slice(0, -1).join(', ') + ' dan ' + proj.selectedDims.slice(-1) : proj.selectedDims[0];
            let saran = cat === 'Sangat Baik' ? 'Terus pertahankan inisiatif ini.' : cat === 'Baik' ? 'Kembangkan terus potensi ini.' : cat === 'Cukup' ? 'Perlu ditingkatkan fokusnya.' : 'Sangat disarankan lebih aktif bertanya.';
            narasiContainer.innerHTML = `Ananda <strong>${sName}</strong> menunjukkan profil <strong>${dimStr}</strong> yang <strong>${cat.toLowerCase()}</strong> dalam kegiatan ${proj.p}. ${proj.desc} ${saran}`; narasiContainer.classList.remove('text-muted');
        }

        // Membuka kembali data kokurikuler yang sudah tersimpan agar bisa dilihat/diubah,
        // dengan mengisi ulang form & tabel input memakai nilai yang sudah ada (bukan tabel kosong).
        function editKokuRecord(id) {
            const k = db.koku.find(x => x.RecordID === id);
            if (!k) return;

            document.getElementById('koku-kelas').value = k.ClassID;
            document.getElementById('koku-tanggal').value = k.Tanggal || '';
            document.getElementById('koku-jenis').value = k.Jenis || 'LintasMapel';
            document.getElementById('koku-projek').value = k.Projek || '';
            document.getElementById('koku-deskripsi').value = k.Deskripsi || '';

            let selectedDims = []; try { selectedDims = JSON.parse(k.Dimensi || '[]'); } catch(e) {}
            document.querySelectorAll('.dim-checkbox').forEach(chk => { chk.checked = selectedDims.includes(chk.value); });

            let kokuData = {}; try { kokuData = JSON.parse(k.KokuData || '{}'); } catch(e) {}

            const stds = db.students.filter(s => String(s.ClassID) === String(k.ClassID));
            window.activeKokuProject = { p: k.Projek, desc: k.Deskripsi, selectedDims };
            let html = '';
            stds.forEach((s, i) => {
                const saved = kokuData[s.StudentID];
                const cat = saved ? saved.Kategori : '';
                const narasi = saved ? saved.Narasi : 'Pilih kategori...';
                html += `<tr data-sid="${s.StudentID}" class="koku-row"><td>${i+1}</td><td class="fw-bold">${s.StudentName}</td><td><select class="form-select form-select-sm koku-cat-select" onchange="updateKokuNarasi(this)">
                    <option value="">Pilih</option>
                    <option value="Sangat Baik" ${cat==='Sangat Baik'?'selected':''}>Sangat Baik</option>
                    <option value="Baik" ${cat==='Baik'?'selected':''}>Baik</option>
                    <option value="Cukup" ${cat==='Cukup'?'selected':''}>Cukup</option>
                    <option value="Kurang" ${cat==='Kurang'?'selected':''}>Kurang</option>
                </select></td><td><small class="koku-narasi-text ${cat ? '' : 'text-muted'}">${narasi}</small></td></tr>`;
            });
            document.getElementById('koku-tbody').innerHTML = html;
            document.getElementById('koku-result-section').style.display = 'block';
            editingKokuId = id;
            window.scrollTo({top: 0, behavior: 'smooth'});
            Swal.fire({icon: 'info', title: 'Mode Edit', text: `Anda sedang membuka data "${k.Projek}". Ubah nilainya lalu klik "Simpan Data" untuk memperbarui.`, timer: 2500, showConfirmButton: false});
        }

        function deleteKokuRecord(id) {
            Swal.fire({
                title: 'Hapus Data Kokurikuler?', text: 'Seluruh nilai kokurikuler pada projek ini akan dihapus.', icon: 'warning',
                showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    db.koku = db.koku.filter(k => k.RecordID !== id);
                    if (editingKokuId === id) { editingKokuId = null; document.getElementById('koku-result-section').style.display = 'none'; }
                    syncDB(); loadKokuHistory();
                    Swal.fire('Terhapus!', 'Data kokurikuler berhasil dihapus.', 'success');
                }
            });
        }

        function saveAllKoku() {
            const cid = document.getElementById('koku-kelas').value;
            const tanggal = document.getElementById('koku-tanggal').value;
            const jenis = document.getElementById('koku-jenis').value;
            const proj = window.activeKokuProject;
            if (!proj) return Swal.fire('Error', 'Silakan buat format tabel terlebih dahulu.', 'error');

            let kokuData = {}, count = 0;
            document.querySelectorAll('.koku-row').forEach(row => {
                const sid = row.getAttribute('data-sid'), cat = row.querySelector('.koku-cat-select').value;
                if (cat) { kokuData[sid] = { Kategori: cat, Narasi: row.querySelector('.koku-narasi-text').innerHTML }; count++; }
            });
            if (count === 0) return Swal.fire('Peringatan', 'Belum ada yang dipilih!', 'warning');

            const record = {
                RecordID: editingKokuId || ("K" + Date.now()),
                ClassID: cid, Tanggal: tanggal, Jenis: jenis,
                Projek: proj.p, Deskripsi: proj.desc, Dimensi: JSON.stringify(proj.selectedDims),
                KokuData: JSON.stringify(kokuData)
            };

            if (editingKokuId) {
                const idx = db.koku.findIndex(k => k.RecordID === editingKokuId);
                if (idx > -1) db.koku[idx] = record; else db.koku.push(record);
            } else {
                db.koku.push(record);
            }

            editingKokuId = null;
            syncDB(); loadKokuHistory();
            Swal.fire('Sukses', `${count} Nilai kokurikuler berhasil disimpan.`, 'success');
        }

        // Mengambil data kokurikuler kelas terpilih, dipersempit sesuai filter projek yang dipilih
        // user (kosong = semua projek kelas ini, atau satu RecordID = hanya projek itu).
        function getKokuExportData() {
            const cid = document.getElementById('koku-kelas').value;
            if (!cid) { Swal.fire('Error', 'Pilih kelas terlebih dahulu.', 'error'); return null; }
            const filterId = document.getElementById('koku-export-filter').value;
            let records = db.koku.filter(k => String(k.ClassID) === String(cid));
            if (filterId) records = records.filter(k => k.RecordID === filterId);
            if (records.length === 0) { Swal.fire('Info', 'Belum ada data kokurikuler untuk diunduh.', 'info'); return null; }

            const className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
            const fileSuffix = filterId ? `_${records[0].Projek.replace(/[^a-z0-9]+/gi, '_')}` : '';
            return { records, className, filterId, fileSuffix };
        }

        async function downloadKokuPDF() {
            const data = getKokuExportData();
            if (!data) return;
            const { records, className, filterId, fileSuffix } = data;

            await ensurePDF();
            const { jsPDF } = window.jspdf; const doc = new jsPDF();
            doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.text(`Rekap Penilaian Kokurikuler - ${className}`, 14, 16);
            let startY = 22;
            if (filterId) { doc.setFontSize(10); doc.setFont("helvetica", "normal"); doc.text(`Projek: ${records[0].Projek}${records[0].Tanggal ? ' (' + records[0].Tanggal + ')' : ''}`, 14, 22); startY = 28; }

            let body = [];
            records.forEach(k => {
                let kokuData = {}; try { kokuData = JSON.parse(k.KokuData || '{}'); } catch(e) {}
                Object.keys(kokuData).forEach(sid => {
                    const student = db.students.find(s => s.StudentID === sid);
                    const name = student ? student.StudentName : sid;
                    body.push([k.Projek, name, kokuData[sid].Narasi.replace(/<[^>]+>/g, '')]);
                });
            });

            doc.autoTable({ head: [['Projek', 'Nama', 'Deskripsi']], body, startY, styles: { fontSize: 9 } });
            doc.save(`Rekap_Kokurikuler_${className}${fileSuffix}.pdf`);
        }

        // Export ke Word (.doc) memakai trik HTML-sebagai-dokumen-Word, sehingga bisa dibuka
        // langsung oleh Microsoft Word/WPS tanpa perlu library docx tambahan di browser.
        function downloadKokuWord() {
            const data = getKokuExportData();
            if (!data) return;
            const { records, className, fileSuffix } = data;

            let bodyHtml = `<h2 style="font-family:Calibri,Arial,sans-serif;">Rekap Penilaian Kokurikuler - ${escapeHtml(className)}</h2>`;
            records.forEach(k => {
                let kokuData = {}; try { kokuData = JSON.parse(k.KokuData || '{}'); } catch(e) {}
                bodyHtml += `<h3 style="font-family:Calibri,Arial,sans-serif;">${escapeHtml(k.Projek || '-')}${k.Tanggal ? ' (' + k.Tanggal + ')' : ''}</h3>`;
                if (k.Deskripsi) bodyHtml += `<p style="font-family:Calibri,Arial,sans-serif;"><i>${escapeHtml(k.Deskripsi)}</i></p>`;
                bodyHtml += `<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; width:100%; font-family:Calibri,Arial,sans-serif; font-size:11pt;">
                    <tr style="background:#f4f7f6;"><th style="width:25%;">Nama Siswa</th><th style="width:15%;">Kategori</th><th>Deskripsi</th></tr>`;
                Object.keys(kokuData).forEach(sid => {
                    const student = db.students.find(s => s.StudentID === sid);
                    const name = student ? student.StudentName : sid;
                    bodyHtml += `<tr><td>${escapeHtml(name)}</td><td>${escapeHtml(kokuData[sid].Kategori)}</td><td>${kokuData[sid].Narasi.replace(/<[^>]+>/g, '')}</td></tr>`;
                });
                bodyHtml += `</table><br>`;
            });

            const htmlDoc = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
                <head><meta charset="utf-8"><title>Rekap Kokurikuler</title></head>
                <body>${bodyHtml}</body></html>`;

            const blob = new Blob(['﻿', htmlDoc], { type: 'application/msword' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `Rekap_Kokurikuler_${className}${fileSuffix}.doc`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);
        }

        // --- Jurnal Harian ---
        function renderJurnal() {
            const classFilter = document.getElementById('filter-jurnal-kelas').value, tbody = document.getElementById('jurnal-tbody');
            if (!classFilter) { tbody.innerHTML = '<tr><td colspan="6" class="text-center">Pilih kelas dahulu</td></tr>'; return; }
            let filtered = db.journal.filter(j => String(j.ClassID) === String(classFilter) || j.ClassID === 'All');
            if (filtered.length === 0) { tbody.innerHTML = '<tr><td colspan="6" class="text-center">Belum ada jurnal</td></tr>'; return; }
            let html = ''; filtered.forEach(j => {
                let attachHtml = j.AttachmentUrl ? `<a href="${j.AttachmentUrl}" target="_blank" class="btn btn-sm btn-info text-white"><i class="bi bi-link-45deg"></i> Buka</a>` : '-';
                html += `<tr><td class="fw-bold">${j.Date}</td><td><span class="badge bg-info text-dark">${j.ClassID==='All'?'Semua Kelas':db.classes.find(c=>String(c.ClassID)===String(j.ClassID))?.ClassName||j.ClassID}</span></td><td class="fw-semibold text-primary">${j.Subject}</td><td><small>${j.Content}</small></td><td>${attachHtml}</td>
                <td class="text-nowrap"><button class="btn btn-sm btn-primary me-1 mb-1" onclick="editJournal('${j.RecordID}')" title="Edit Jurnal"><i class="bi bi-pencil-square"></i></button><button class="btn btn-sm btn-danger mb-1" onclick="deleteJournal('${j.RecordID}')" title="Hapus Jurnal"><i class="bi bi-trash"></i></button></td></tr>`;
            });
            tbody.innerHTML = html;
        }

        function saveJurnal() {
            const d = document.getElementById('jurnal-date').value, k = document.getElementById('jurnal-kelas').value, m = document.getElementById('jurnal-mapel').value, c = document.getElementById('jurnal-content').value;
            const fileInput = document.getElementById('jurnal-file');
            const syncCal = document.getElementById('jurnal-sync-cal').checked;
            if(!d || !k || !m || !c) return Swal.fire('Error', 'Harap isi semua field utama!', 'error');

            Swal.fire({title: 'Menyimpan...', allowOutsideClick: false, didOpen: () => { Swal.showLoading(); }});

            if(fileInput.files.length > 0) {
                let file = fileInput.files[0];
                let reader = new FileReader();
                reader.onload = function(e) {
                    if(typeof google !== 'undefined') {
                        google.script.run.withSuccessHandler(function(url) {
                            finishSaveJurnal(d, k, m, c, url, syncCal);
                        }).withFailureHandler(function() {
                            finishSaveJurnal(d, k, m, c, "", syncCal);
                        }).uploadFileToDrive(e.target.result, file.name, file.type);
                    } else {
                        // Simulasi URL saat tidak di dalam GAS
                        finishSaveJurnal(d, k, m, c, "https://simulasi-url-drive.com/file", syncCal);
                    }
                };
                reader.readAsDataURL(file);
            } else {
                // null = tidak ada file baru yang dipilih; saat mode edit, lampiran lama tetap dipertahankan
                finishSaveJurnal(d, k, m, c, null, syncCal);
            }
        }

        function finishSaveJurnal(d, k, m, c, fileUrl, syncCal) {
            let className = db.classes.find(cls => String(cls.ClassID) === String(k))?.ClassName || k;
            if(syncCal && typeof google !== 'undefined') {
                google.script.run.syncToCalendar('jurnal', {date: d, subject: m, className: className, content: c});
            }

            if (editingJournalId) {
                const idx = db.journal.findIndex(j => j.RecordID === editingJournalId);
                if (idx > -1) {
                    db.journal[idx].Date = d; db.journal[idx].ClassID = k; db.journal[idx].Subject = m; db.journal[idx].Content = c;
                    if (fileUrl !== null) db.journal[idx].AttachmentUrl = fileUrl; // null = pertahankan lampiran lama
                }
                cancelEditJurnal();
                Swal.fire('Tersimpan', 'Jurnal berhasil diperbarui.', 'success');
            } else {
                db.journal.push({ RecordID: "J" + Date.now(), Date: d, ClassID: k, Subject: m, Content: c, AttachmentUrl: fileUrl || "" });
                Swal.fire('Tersimpan', 'Jurnal disimpan.', 'success');
            }

            document.getElementById('filter-jurnal-kelas').value = k; renderJurnal(); syncDB();
            document.getElementById('jurnal-content').value = ''; document.getElementById('jurnal-file').value = '';
        }

        function editJournal(id) {
            const j = db.journal.find(x => x.RecordID === id);
            if (!j) return;
            editingJournalId = id;
            document.getElementById('jurnal-date').value = j.Date;
            document.getElementById('jurnal-kelas').value = j.ClassID;
            document.getElementById('jurnal-mapel').value = j.Subject;
            document.getElementById('jurnal-content').value = j.Content;
            document.getElementById('jurnal-file').value = '';
            document.getElementById('btn-save-jurnal').innerHTML = '<i class="bi bi-save me-1"></i> Update Jurnal';
            document.getElementById('btn-cancel-edit-jurnal').style.display = 'block';
            window.scrollTo({top: 0, behavior: 'smooth'});
        }

        function cancelEditJurnal() {
            editingJournalId = null;
            document.getElementById('btn-save-jurnal').innerHTML = '<i class="bi bi-save me-1"></i> Simpan';
            document.getElementById('btn-cancel-edit-jurnal').style.display = 'none';
            document.getElementById('jurnal-content').value = '';
            document.getElementById('jurnal-file').value = '';
        }

        function deleteJournal(id) {
            Swal.fire({
                title: 'Hapus Jurnal?', text: "Apakah Anda yakin ingin menghapus catatan jurnal ini?", icon: 'warning',
                showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Ya, Hapus!'
            }).then((result) => {
                if (result.isConfirmed) {
                    db.journal = db.journal.filter(j => j.RecordID !== id);
                    if (editingJournalId === id) cancelEditJurnal();
                    syncDB(); renderJurnal();
                    Swal.fire('Terhapus!', 'Jurnal berhasil dihapus.', 'success');
                }
            });
        }

        async function downloadJurnalPDF() {
            const cFilter = document.getElementById('filter-jurnal-kelas').value;
            if (!cFilter) return Swal.fire('Peringatan', 'Pilih kelas', 'warning');
            let filtered = db.journal.filter(j => String(j.ClassID) === String(cFilter) || j.ClassID === 'All');
            if (filtered.length === 0) return Swal.fire('Info', 'Kosong', 'info');
            const className = cFilter === 'All' ? 'Semua Kelas' : (db.classes.find(c => String(c.ClassID) === String(cFilter))?.ClassName || cFilter);
            await ensurePDF();
            const { jsPDF } = window.jspdf; const doc = new jsPDF();
            doc.setFontSize(14); doc.setFont("helvetica", "bold"); doc.text(`Jurnal Harian - ${className}`, 14, 16);
            doc.autoTable({ head: [['Tanggal', 'Kelas', 'Mapel', 'Jurnal']], body: filtered.map(j => [j.Date, j.ClassID, j.Subject, j.Content]), startY: 22, styles: { fontSize: 10 } }); doc.save(`Jurnal.pdf`);
        }

        // --- Profile ---
        let tempPhotoBase64 = "";
        function updateProfileUI() {
            if (!db.config.userName) db.config.userName = "Resti Novika, M.Pd."; if (!db.config.userSchool) db.config.userSchool = "Guru Kelas";
            document.getElementById('profile-name').innerText = db.config.userName; document.getElementById('profile-school').innerText = db.config.userSchool;
            document.getElementById('profile-img').src = db.config.userPhoto || ('https://ui-avatars.com/api/?name=' + encodeURIComponent(db.config.userName) + '&background=FF7F00&color=fff');
        }

        function openProfileModal() {
            document.getElementById('prof-name').value = db.config.userName || ''; document.getElementById('prof-school').value = db.config.userSchool || '';
            document.getElementById('modal-prof-img').src = db.config.userPhoto || ('https://ui-avatars.com/api/?name=' + encodeURIComponent(db.config.userName) + '&background=FF7F00&color=fff');
            tempPhotoBase64 = db.config.userPhoto || ""; document.getElementById('prof-photo-file').value = ""; 
            let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalProfile'));
            if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalProfile')); modalInstance.show();
        }

        function closeModal(id) { let modalInstance = bootstrap.Modal.getInstance(document.getElementById(id)); if (modalInstance) modalInstance.hide(); }

        function previewProfilePhoto(event) {
            const file = event.target.files[0];
            if (file) { const r = new FileReader(); r.onload = function(e) { const img = new Image(); img.onload = function() {
                        const canvas = document.createElement('canvas'); const ctx = canvas.getContext('2d'); canvas.width = 300; canvas.height = 300; ctx.drawImage(img, 0, 0, 300, 300);
                        tempPhotoBase64 = canvas.toDataURL('image/jpeg', 0.7); document.getElementById('modal-prof-img').src = tempPhotoBase64;
                    }; img.src = e.target.result; }; r.readAsDataURL(file); }
        }

        function saveProfile() {
            db.config.userName = document.getElementById('prof-name').value || 'Guru'; db.config.userSchool = document.getElementById('prof-school').value || 'Sekolah';
            if (tempPhotoBase64) db.config.userPhoto = tempPhotoBase64;
            syncDB(); updateProfileUI(); closeModal('modalProfile'); Swal.fire('Tersimpan', 'Profil diperbarui!', 'success');
        }

        // --- Schedule / Jadwal ---
        function renderSchedule() {
            const cFilter = document.getElementById('filter-sched-class').value, container = document.getElementById('schedule-container');
            if (!cFilter) { container.innerHTML = `<div class="col-12 text-center text-muted py-5"><i class="bi bi-calendar-x" style="font-size: 3rem;"></i><p>Pilih kelas terlebih dahulu.</p></div>`; return; }
            let html = '';
            ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'].forEach(day => {
                let daySch = db.schedule.filter(s => s.Day === day && (String(s.ClassID) === String(cFilter) || s.ClassID === 'All')).sort((a,b) => a.TimeStart.localeCompare(b.TimeStart));
                html += `<div class="col-md-4 mb-4"><div class="card shadow-sm border-0 h-100"><div class="card-header bg-primary text-white fw-bold text-center py-2">${day}</div><ul class="list-group list-group-flush">`;
                if(daySch.length === 0) {
                    html += `<li class="list-group-item text-center text-muted small py-3">Libur / Kosong</li>`;
                } else {
                    daySch.forEach(s => { 
                        let className = s.ClassID === 'All' ? 'Semua Kelas' : (db.classes.find(c => String(c.ClassID) === String(s.ClassID))?.ClassName || s.ClassID);
                        html += `
                        <li class="list-group-item px-3 py-2">
                            <div class="d-flex justify-content-between align-items-center mb-1">
                                <span class="badge bg-warning text-dark"><i class="bi bi-clock"></i> ${s.TimeStart} - ${s.TimeEnd}</span>
                                <div>
                                    <button class="btn btn-sm btn-link text-primary p-0 me-2" onclick="editSchedule('${s.RecordID}')" title="Edit Jadwal"><i class="bi bi-pencil-square"></i></button>
                                    <button class="btn btn-sm btn-link text-danger p-0" onclick="deleteSchedule('${s.RecordID}')" title="Hapus Jadwal"><i class="bi bi-trash"></i></button>
                                </div>
                            </div>
                            <div class="fw-semibold text-primary">${s.Subject}</div>
                            <small class="text-muted fw-bold" style="font-size:0.75rem">${className}</small>
                        </li>`; 
                    });
                }
                html += `</ul></div></div>`;
            }); container.innerHTML = html;
        }

        function showAddScheduleModal() { 
            editingScheduleId = null; 
            document.getElementById('modal-schedule-title').innerText = "Tambah Jadwal Pelajaran";
            document.getElementById('sched-day').value = 'Senin';
            document.getElementById('sched-start').value = '';
            document.getElementById('sched-end').value = '';
            document.getElementById('sched-subject').value = '';
            document.getElementById('sched-class').value = '';
            
            let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalAddSchedule'));
            if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalAddSchedule')); 
            modalInstance.show(); 
        }

        function editSchedule(id) {
            const s = db.schedule.find(x => x.RecordID === id);
            if (!s) return;
            
            editingScheduleId = id;
            document.getElementById('modal-schedule-title').innerText = "Edit Jadwal Pelajaran";
            document.getElementById('sched-day').value = s.Day;
            document.getElementById('sched-start').value = s.TimeStart;
            document.getElementById('sched-end').value = s.TimeEnd;
            document.getElementById('sched-subject').value = s.Subject;
            document.getElementById('sched-class').value = s.ClassID;

            let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalAddSchedule'));
            if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalAddSchedule')); 
            modalInstance.show();
        }

        function deleteSchedule(id) {
            Swal.fire({
                title: 'Hapus Jadwal?', text: "Apakah Anda yakin ingin menghapus jadwal ini?", icon: 'warning',
                showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Ya, Hapus!'
            }).then((result) => {
                if (result.isConfirmed) {
                    db.schedule = db.schedule.filter(s => s.RecordID !== id);
                    syncDB(); renderSchedule(); Swal.fire('Terhapus!', 'Jadwal berhasil dihapus.', 'success');
                }
            });
        }

        function saveSchedule() {
            const day = document.getElementById('sched-day').value, 
                  start = document.getElementById('sched-start').value, 
                  end = document.getElementById('sched-end').value, 
                  subj = document.getElementById('sched-subject').value, 
                  cid = document.getElementById('sched-class').value;
            const syncCal = document.getElementById('sched-sync-cal').checked;
            
            if(!start || !end || !subj || !cid) return Swal.fire('Error', 'Harap lengkapi semua form!', 'error');
            
            let className = db.classes.find(c => String(c.ClassID) === String(cid))?.ClassName || cid;
            
            if (editingScheduleId) {
                let idx = db.schedule.findIndex(s => s.RecordID === editingScheduleId);
                if (idx >= 0) {
                    db.schedule[idx].Day = day; db.schedule[idx].TimeStart = start; db.schedule[idx].TimeEnd = end; db.schedule[idx].Subject = subj; db.schedule[idx].ClassID = cid;
                }
            } else {
                db.schedule.push({ RecordID: "SCH" + Date.now(), Day: day, TimeStart: start, TimeEnd: end, Subject: subj, ClassID: cid });
                
                // Tambah ke Calendar otomatis khusus saat Jadwal Baru dibuat (bukan edit)
                if(syncCal && typeof google !== 'undefined') {
                    google.script.run.syncToCalendar('schedule', {day: day, start: start, end: end, subject: subj, className: className});
                }
            }
            
            closeModal('modalAddSchedule'); renderSchedule(); syncDB(); 
            let msg = editingScheduleId ? 'Jadwal diperbarui.' : 'Jadwal ditambahkan.';
            Swal.fire('Tersimpan', msg, 'success');
            editingScheduleId = null;
        }

        // --- Poin Siswa ---
        function populatePoinStudentSelect() {
            const cid = document.getElementById('poin-kelas').value;
            const sel = document.getElementById('poin-siswa');
            if (!cid) { sel.innerHTML = '<option value="">-- Pilih Kelas Dahulu --</option>'; return; }

            const stds = cid === 'All' ? db.students : db.students.filter(s => String(s.ClassID) === String(cid));
            let opts = '<option value="">-- Pilih Siswa --</option>';
            stds.forEach(s => opts += `<option value="${s.StudentID}">${s.StudentName}</option>`);
            sel.innerHTML = opts;
        }

        function savePoin() {
            const sid = document.getElementById('poin-siswa').value;
            const tanggal = document.getElementById('poin-tanggal').value;
            const reward = parseInt(document.getElementById('poin-reward').value) || 0;
            const pelanggaran = parseInt(document.getElementById('poin-pelanggaran').value) || 0;
            const keterangan = document.getElementById('poin-keterangan').value.trim();

            if (!sid) return Swal.fire('Error', 'Harap pilih siswa terlebih dahulu.', 'error');
            if (!tanggal) return Swal.fire('Error', 'Harap isi tanggal.', 'error');
            if (reward === 0 && pelanggaran === 0) return Swal.fire('Error', 'Isi minimal salah satu Poin Reward atau Poin Pelanggaran.', 'error');

            const student = db.students.find(s => s.StudentID === sid);
            if (!student) return Swal.fire('Error', 'Data siswa tidak ditemukan.', 'error');

            const record = {
                RecordID: editingPoinId || ("P" + Date.now()),
                StudentID: sid, ClassID: student.ClassID, Tanggal: tanggal,
                PoinReward: reward, PoinPelanggaran: pelanggaran, Keterangan: keterangan
            };

            if (editingPoinId) {
                const idx = db.poin.findIndex(p => p.RecordID === editingPoinId);
                if (idx > -1) db.poin[idx] = record; else db.poin.push(record);
            } else {
                db.poin.push(record);
            }

            let msg = editingPoinId ? 'Poin siswa berhasil diperbarui.' : 'Poin siswa berhasil disimpan.';
            cancelEditPoin();
            document.getElementById('filter-poin-kelas').value = student.ClassID;
            syncDB(); loadPoinData();
            Swal.fire('Tersimpan', msg, 'success');
        }

        function editPoin(id) {
            const p = db.poin.find(x => x.RecordID === id);
            if (!p) return;
            editingPoinId = id;
            document.getElementById('poin-kelas').value = p.ClassID;
            populatePoinStudentSelect();
            document.getElementById('poin-siswa').value = p.StudentID;
            document.getElementById('poin-tanggal').value = p.Tanggal;
            document.getElementById('poin-reward').value = p.PoinReward;
            document.getElementById('poin-pelanggaran').value = p.PoinPelanggaran;
            document.getElementById('poin-keterangan').value = p.Keterangan || '';
            document.getElementById('btn-save-poin').innerHTML = '<i class="bi bi-save me-1"></i> Update Poin';
            document.getElementById('btn-cancel-edit-poin').style.display = 'block';
            window.scrollTo({top: 0, behavior: 'smooth'});
        }

        function cancelEditPoin() {
            editingPoinId = null;
            document.getElementById('btn-save-poin').innerHTML = '<i class="bi bi-save me-1"></i> Simpan';
            document.getElementById('btn-cancel-edit-poin').style.display = 'none';
            document.getElementById('poin-reward').value = 0;
            document.getElementById('poin-pelanggaran').value = 0;
            document.getElementById('poin-keterangan').value = '';
        }

        function deletePoin(id) {
            Swal.fire({
                title: 'Hapus Data Poin?', text: 'Data poin siswa ini akan dihapus secara permanen.', icon: 'warning',
                showCancelButton: true, confirmButtonColor: '#d33', confirmButtonText: 'Ya, Hapus!', cancelButtonText: 'Batal'
            }).then((result) => {
                if (result.isConfirmed) {
                    db.poin = db.poin.filter(p => p.RecordID !== id);
                    if (editingPoinId === id) cancelEditPoin();
                    syncDB(); loadPoinData();
                    Swal.fire('Terhapus!', 'Data poin berhasil dihapus.', 'success');
                }
            });
        }

        function loadPoinData() {
            const cid = document.getElementById('filter-poin-kelas').value;
            const tbody = document.getElementById('poin-tbody'), rekapTbody = document.getElementById('poin-rekap-tbody');
            if (!cid) {
                tbody.innerHTML = '<tr><td colspan="8" class="text-center">Pilih kelas dahulu</td></tr>';
                rekapTbody.innerHTML = '<tr><td colspan="4" class="text-center">Pilih kelas dahulu</td></tr>';
                return;
            }

            const records = db.poin.filter(p => cid === 'All' || String(p.ClassID) === String(cid))
                .slice().sort((a, b) => String(b.Tanggal).localeCompare(String(a.Tanggal)));

            if (records.length === 0) {
                tbody.innerHTML = '<tr><td colspan="8" class="text-center">Belum ada data poin</td></tr>';
                rekapTbody.innerHTML = '<tr><td colspan="4" class="text-center">Belum ada data poin</td></tr>';
                return;
            }

            let html = '';
            records.forEach(p => {
                const student = db.students.find(s => s.StudentID === p.StudentID);
                const name = student ? student.StudentName : '(Siswa tidak ditemukan)';
                const className = db.classes.find(c => String(c.ClassID) === String(p.ClassID))?.ClassName || p.ClassID;
                // PENTING: paksa jadi Number() karena setelah round-trip lewat Spreadsheet (kolom
                // diformat Teks), nilai ini bisa datang sebagai string "5" bukan angka 5 — kalau
                // langsung dijumlahkan dengan "+=" akan jadi penggabungan string ("05"), bukan hitung.
                const reward = Number(p.PoinReward) || 0, pelanggaran = Number(p.PoinPelanggaran) || 0;
                const total = reward - pelanggaran;
                html += `<tr><td>${p.Tanggal || '-'}</td><td class="fw-bold">${escapeHtml(name)}</td><td>${escapeHtml(className)}</td>
                    <td class="text-success fw-bold">+${reward}</td><td class="text-danger fw-bold">-${pelanggaran}</td>
                    <td class="fw-bold ${total >= 0 ? 'text-success' : 'text-danger'}">${total >= 0 ? '+' : ''}${total}</td>
                    <td><small>${escapeHtml(p.Keterangan || '-')}</small></td>
                    <td class="text-nowrap"><button class="btn btn-sm btn-primary me-1 mb-1" onclick="editPoin('${p.RecordID}')" title="Edit"><i class="bi bi-pencil-square"></i></button><button class="btn btn-sm btn-danger mb-1" onclick="deletePoin('${p.RecordID}')" title="Hapus"><i class="bi bi-trash"></i></button></td></tr>`;
            });
            tbody.innerHTML = html;

            // Rekap total akumulasi poin per siswa
            let rekap = {};
            records.forEach(p => {
                if (!rekap[p.StudentID]) rekap[p.StudentID] = { reward: 0, pelanggaran: 0 };
                rekap[p.StudentID].reward += Number(p.PoinReward) || 0;
                rekap[p.StudentID].pelanggaran += Number(p.PoinPelanggaran) || 0;
            });
            let rekapHtml = '';
            Object.keys(rekap).sort((a, b) => {
                const nameA = db.students.find(s => s.StudentID === a)?.StudentName || '';
                const nameB = db.students.find(s => s.StudentID === b)?.StudentName || '';
                return nameA.localeCompare(nameB);
            }).forEach(sid => {
                const student = db.students.find(s => s.StudentID === sid);
                const name = student ? student.StudentName : '(Siswa tidak ditemukan)';
                const r = rekap[sid], total = r.reward - r.pelanggaran;
                rekapHtml += `<tr><td class="fw-bold">${escapeHtml(name)}</td><td class="text-success">+${r.reward}</td><td class="text-danger">-${r.pelanggaran}</td><td class="fw-bold ${total >= 0 ? 'text-success' : 'text-danger'}">${total >= 0 ? '+' : ''}${total}</td></tr>`;
            });
            rekapTbody.innerHTML = rekapHtml;
        }

        // =====================================
        // INTEGRASI WHATSAPP HELPER
        // =====================================
        function sendWAAbsen(sid) {
            let student = db.students.find(s => s.StudentID === sid);
            if(!student || !student.ParentPhone) return Swal.fire('Gagal', 'Nomor HP Orang Tua belum diatur untuk siswa ini.', 'error');
            
            // Perbaikan Bug: Membersihkan karakter selain angka (seperti spasi atau strip)
            let phone = student.ParentPhone.replace(/\D/g, ''); 
            if (phone.startsWith('0')) phone = phone.replace(/^0/, '62'); 
            
            let selectElem = document.getElementById('absen-bulan');
            let month = selectElem.options[selectElem.selectedIndex].text;
            
            // Hitung rekap
            let tH=0, tS=0, tI=0, tA=0;
            document.querySelectorAll(`.select-absen-val[data-sid="${sid}"]`).forEach(sel => {
                let v = sel.value;
                if(v==='H') tH++; else if(v==='S') tS++; else if(v==='I') tI++; else if(v==='A') tA++;
            });
            
            let guruName = db.config.userName || 'Guru Kelas';
            let text = `Halo Bapak/Ibu Wali Murid dari *${student.StudentName}*,\n\nBerikut adalah rekap sementara kehadiran ananda pada bulan *${month}*:\n\n✔️ Hadir (H): ${tH}\n🤒 Sakit (S): ${tS}\n📩 Izin (I): ${tI}\n❌ Alpha (A): ${tA}\n\nMohon bantuannya untuk memantau kehadiran ananda. Terima kasih,\n*${guruName}*`;
            
            // Perbaikan Bug: Menggunakan metode pembuatan elemen anchor (<a>) dinamis untuk menghindari pemblokiran Pop-up dari Iframe
            let waUrl = 'https://api.whatsapp.com/send?phone=' + phone + '&text=' + encodeURIComponent(text);
            let a = document.createElement('a');
            a.target = '_blank';
            a.href = waUrl;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }

        function sendWANilai(sid) {
            let student = db.students.find(s => s.StudentID === sid);
            if(!student || !student.ParentPhone) return Swal.fire('Gagal', 'Nomor HP Orang Tua belum diatur untuk siswa ini.', 'error');
            
            // Perbaikan Bug
            let phone = student.ParentPhone.replace(/\D/g, '');
            if (phone.startsWith('0')) phone = phone.replace(/^0/, '62');
            
            let mapel = document.getElementById('nilai-mapel').value;
            let selJenis = document.getElementById('nilai-jenis');
            let jenis = selJenis.options[selJenis.selectedIndex].text;
            
            let vals = [];
            document.querySelectorAll(`.input-nilai-val[data-sid="${sid}"]`).forEach(inp => vals.push(inp.value || '-'));
            
            let guruName = db.config.userName || 'Guru Kelas';
            let text = `Halo Bapak/Ibu Wali Murid dari *${student.StudentName}*,\n\nBerikut adalah hasil penilaian *${jenis}* untuk mata pelajaran *${mapel}*:\n\n*Nilai: ${vals.join(', ')}*\n\nTerima kasih atas bimbingan Bapak/Ibu di rumah,\n*${guruName}*`;
            
            // Perbaikan Bug
            let waUrl = 'https://api.whatsapp.com/send?phone=' + phone + '&text=' + encodeURIComponent(text);
            let a = document.createElement('a');
            a.target = '_blank';
            a.href = waUrl;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }

        // =====================================
        // FITUR BROADCAST WHATSAPP (UMUM - MANAJEMEN KELAS SAJA)
        // =====================================
        let broadcastQueue = [];
        let currentBroadcastIndex = 0;

        function toggleAllStudents(source) {
            const checkboxes = document.querySelectorAll('.chk-std:not(:disabled)');
            checkboxes.forEach(chk => chk.checked = source.checked);
        }

        function openBroadcastModal() {
            let selectedChks = document.querySelectorAll('.chk-std:checked');

            if(!selectedChks || selectedChks.length === 0) {
                return Swal.fire('Info', 'Pilih minimal 1 siswa dengan mencentang kotak di tabel. (Pastikan siswa tersebut memiliki No HP)', 'info');
            }

            broadcastQueue = [];
            let htmlList = '';

            selectedChks.forEach(chk => {
                let student = db.students.find(s => s.StudentID === chk.value);
                if(student && student.ParentPhone) {
                    broadcastQueue.push(student);
                    htmlList += `<li class="list-group-item px-3 py-2 small d-flex align-items-center" id="br-item-${student.StudentID}"><i class="bi bi-clock-history text-secondary me-2 fs-5"></i> <div><span class="fw-bold">${student.StudentName}</span><br><small class="text-muted">${student.ParentPhone}</small></div></li>`;
                }
            });

            if(broadcastQueue.length === 0) return Swal.fire('Info', 'Siswa yang dipilih tidak memiliki No HP.', 'info');

            document.getElementById('broadcast-count').innerText = broadcastQueue.length;
            document.getElementById('broadcast-list').innerHTML = htmlList;
            
            document.getElementById('broadcast-msg-wrapper').style.display = 'block';
            document.getElementById('broadcast-msg').value = '';
            
            currentBroadcastIndex = 0;
            const btnSend = document.getElementById('btn-send-broadcast');
            btnSend.innerText = `Kirim Pesan (1 dari ${broadcastQueue.length})`;
            btnSend.classList.replace('btn-secondary', 'btn-success');
            btnSend.disabled = false;

            let modalInstance = bootstrap.Modal.getInstance(document.getElementById('modalBroadcast'));
            if (!modalInstance) modalInstance = new bootstrap.Modal(document.getElementById('modalBroadcast')); 
            modalInstance.show();
        }

        function sendNextBroadcast() {
            const msgTemplate = document.getElementById('broadcast-msg').value.trim();
            if(!msgTemplate) return Swal.fire('Peringatan', 'Pesan pengumuman tidak boleh kosong!', 'warning');

            if(currentBroadcastIndex >= broadcastQueue.length) {
                Swal.fire('Selesai', 'Semua pesan broadcast telah diproses!', 'success');
                return;
            }

            let student = broadcastQueue[currentBroadcastIndex];
            
            // Ganti placeholder [Nama] dengan nama asli siswa
            let text = msgTemplate.replace(/\[Nama\]/gi, student.StudentName);
            
            let phone = student.ParentPhone.replace(/\D/g, '');
            if (phone.startsWith('0')) phone = phone.replace(/^0/, '62');
            
            // Coret siswa dari antrean karena link WA sudah dibuka
            let liElem = document.getElementById(`br-item-${student.StudentID}`);
            if(liElem) {
                liElem.innerHTML = `<i class="bi bi-check-circle-fill text-success me-2 fs-5"></i> <div><s class="fw-bold">${student.StudentName}</s><br><small class="text-muted">${student.ParentPhone}</small></div>`;
                liElem.classList.add('bg-light');
            }

            let waUrl = 'https://api.whatsapp.com/send?phone=' + phone + '&text=' + encodeURIComponent(text);
            let a = document.createElement('a');
            a.target = '_blank';
            a.href = waUrl;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);

            currentBroadcastIndex++;
            
            const btnSend = document.getElementById('btn-send-broadcast');
            if(currentBroadcastIndex < broadcastQueue.length) {
                btnSend.innerHTML = `Kirim Selanjutnya (${currentBroadcastIndex + 1} dari ${broadcastQueue.length})`;
            } else {
                btnSend.innerHTML = `<i class="bi bi-check2-all"></i> Selesai`;
                btnSend.classList.replace('btn-success', 'btn-secondary');
            }
        }

</script>
