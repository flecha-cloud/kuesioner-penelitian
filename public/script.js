const groups = {
 motivation:[
  ['Kebutuhan fisik',['Saya mendapatkan fasilitas kerja yang memadai dari perusahaan.','Saya merasa kondisi fisik tempat kerja nyaman untuk bekerja.','Saya mendapatkan waktu istirahat yang cukup saat bekerja.']],
  ['Kebutuhan rasa aman dan keselamatan',['Saya merasa aman saat bekerja di lingkungan perusahaan ini.','Saya merasa tenang karena perusahaan peduli pada keselamatan kerja.']],
  ['Kebutuhan sosial',['Saya sering diajak berdiskusi atau kerja sama dalam tim.','Saya memiliki rekan kerja yang membantu ketika saya butuh bantuan.','Saya mendapat dukungan moral dari rekan kerja saat kesulitan.']],
  ['Kebutuhan akan penghargaan',['Atasan saya memberikan apresiasi atas pencapaian saya.','Saya menerima bonus/insetif atas hasil kerja saya.']]
 ],
 discipline:[
  ['Kehadiran kerja',['Saya selalu datang tepat waktu ke tempat kerja.','Saya jarang mengambil cuti tanpa alasan penting.']],
  ['Kepatuhan terhadap aturan perusahaan',['Saya selalu mematuhi aturan kerja yang berlaku.','Saya tidak menggunakan fasilitas kantor untuk keperluan pribadi.']],
  ['Tanggung jawab terhadap pekerjaan',['Saya menyelesaikan pekerjaan sesuai dengan tenggat waktu.','Saya memastikan hasil pekerjaan sesuai standar perusahaan.']],
  ['Konsisten dalam bekerja',['Saya bekerja dengan konsisten dari awal hingga akhir kerja.','Saya mempertahankan kualitas kerja setiap harinya.']],
  ['Sikap terhadap peringatan atau teguran',['Saya menjadikan teguran sebagai bahan perbaikan diri.','Saya tidak merasa tersinggung jika ditegur karena kesalahan.']]
 ],
 performance:[
  ['Kualitas kerja',['Saya berusaha meningkatkan mutu kerja saya secara berkelanjutan.','Saya memastikan hasil kerja saya rapi dan sesuai standar.']],
  ['Kuantitas kerja',['Saya memenuhi target kerja yang diberikan atasan.','Saya tidak pernah menunda-nunda pekerjaan.']],
  ['Tanggung jawab',['Saya merasa bertanggung jawab atas pekerjaan yang saya kerjakan.','Saya bersedia memperbaiki kesalahan dalam pekerjaan saya.']],
  ['Kerja sama',['Saya aktif berkontribusi dalam kegiatan kelompok.','Saya bersedia membantu rekan kerja saat mereka kesulitan.']],
  ['Inisiatif',['Saya mengambil tindakan lebih tanpa harus diperintah.','Saya sering mengusulkan ide-ide baru dalam pekerjaan.']]
 ]
};
const choices=[['STS','Sangat Tidak Setuju'],['TS','Tidak Setuju'],['N','Netral'],['S','Setuju'],['SS','Sangat Setuju']];
const questionTexts=[];let current=0;
function renderQuestions(key,target,start){let n=start;const root=document.getElementById(target);groups[key].forEach((g,idx)=>{const section=document.createElement('section');section.className='question-group';const head=document.createElement('div');head.className='group-head';head.innerHTML=`<span class="group-letter">${String.fromCharCode(65+idx)}</span><span>${g[0]}</span>`;section.append(head);g[1].forEach(q=>{n++;questionTexts[n]=q;const row=document.createElement('div');row.className='question-row';const text=document.createElement('div');text.className='question-text';text.innerHTML=`<span class="question-no">${n}.</span>${q}`;const opts=document.createElement('div');opts.className='likert';choices.forEach(([code,label])=>{const l=document.createElement('label');l.title=label;l.innerHTML=`<input type="radio" name="P${n}" value="${label}"><span>${code}</span><span class="full-label">${label}</span>`;opts.append(l)});row.append(text,opts);section.append(row)});root.append(section)});}
renderQuestions('motivation','motivationQuestions',0);renderQuestions('discipline','disciplineQuestions',10);renderQuestions('performance','performanceQuestions',20);
const pages=[...document.querySelectorAll('.page')],titles=['Identitas Responden','Motivasi Kerja','Disiplin Kerja','Kinerja Pegawai'],descs=['Mohon lengkapi informasi berikut sebelum memulai.','Pernyataan mengenai motivasi kerja. Pilih jawaban yang paling sesuai.','Pernyataan mengenai disiplin kerja. Pilih jawaban yang paling sesuai.','Pernyataan mengenai kinerja pegawai. Pilih jawaban yang paling sesuai.'];
const $=id=>document.getElementById(id);const alertBox=$('alert');
function showAlert(msg){alertBox.textContent=msg;alertBox.hidden=false;alertBox.scrollIntoView({behavior:'smooth',block:'center'})}function clearAlert(){alertBox.hidden=true;alertBox.textContent=''}
function showPage(i){current=i;pages.forEach((p,n)=>p.classList.toggle('active',n===i));$('pageTitle').textContent=titles[i];$('pageDesc').textContent=descs[i];$('pageEyebrow').textContent=`BAGIAN 0${i+1} / 04`;$('pageNumber').textContent=`0${i+1}`;$('stepLabel').textContent=`Langkah ${i+1} dari 4`;$('progressFill').style.width=`${(i+1)*25}%`;$('backBtn').hidden=i===0;$('nextBtn').hidden=i===3;$('submitBtn').hidden=i!==3;clearAlert();window.scrollTo({top:0,behavior:'smooth'});}
function validate(i){if(i===0){if(!document.querySelector('input[name="gender"]:checked')||!$('age').value||!$('education').value){showAlert('Mohon lengkapi semua identitas responden sebelum melanjutkan.');return false}}else{const first=(i-1)*10+1,last=i*10;const missing=[];for(let n=first;n<=last;n++)if(!document.querySelector(`input[name="P${n}"]:checked`))missing.push(n);if(missing.length){showAlert(`Masih ada ${missing.length} pertanyaan yang belum dijawab: ${missing.join(', ')}. Silakan lengkapi terlebih dahulu.`);const el=document.querySelector(`input[name="P${missing[0]}"]`);el?.closest('.question-row')?.scrollIntoView({behavior:'smooth',block:'center'});return false}}return true}
$('nextBtn').addEventListener('click',()=>{if(validate(current))showPage(current+1)});$('backBtn').addEventListener('click',()=>showPage(Math.max(0,current-1)));
let sending=false;function getAnswers(){const data={ 'form-name':'survey-responses','bot-field':'',gender:document.querySelector('input[name="gender"]:checked').value,age:$('age').value,education:$('education').value,submitted_at:new Date().toISOString(),submission_id:(crypto.randomUUID?crypto.randomUUID():`${Date.now()}-${Math.random().toString(16).slice(2)}`)};for(let n=1;n<=30;n++)data[`P${n}`]=document.querySelector(`input[name="P${n}"]:checked`).value;return data}
$('surveyForm').addEventListener('submit',async e=>{e.preventDefault();if(sending)return;if(!validate(3))return;const data=getAnswers();sending=true;$('submitBtn').disabled=true;$('submitBtn').textContent='Mengirim…';clearAlert();try{const body=new URLSearchParams();
Object.entries(data).forEach(([key,value])=>body.append(key,String(value)));
const response=await fetch('/',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'},body:body.toString(),redirect:'follow'});
if(!response.ok)throw new Error(`Server merespons ${response.status}`);
const responseText=await response.text();
if(/page not found|not found|error/i.test(responseText.slice(0,500)) && response.status>=400) throw new Error('Netlify tidak menerima formulir. Periksa Form Detection dan deploy ulang.');$('surveyForm').hidden=true;$('hero').hidden=true;document.querySelector('.page-head').hidden=true;$('successPage').hidden=false;$('stepLabel').textContent='Selesai';$('progressFill').style.width='100%';window.scrollTo({top:0,behavior:'smooth'});}catch(err){console.error(err);showAlert('Pengiriman belum berhasil. Periksa koneksi internet Anda lalu coba kirim kembali. Jawaban tetap tersedia di halaman ini.');sending=false;$('submitBtn').disabled=false;$('submitBtn').innerHTML='Kirim Jawaban <span>✓</span>'}});
$('homeBtn').addEventListener('click',()=>location.reload());
