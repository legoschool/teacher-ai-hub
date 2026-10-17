const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];let all=false;function notify(t){$('.toast').textContent=t;$('.toast').hidden=false;clearTimeout(window.toastTimer);window.toastTimer=setTimeout(()=>$('.toast').hidden=true,2400)}function route(){
 const requested=location.hash.slice(1)||'home';
 const target=document.getElementById(requested);
 const view=target?.closest('.view');
 const id=view?view.id:'home';
 $$('.view').forEach(x=>x.hidden=!all&&x.id!==id);
 document.body.classList.toggle('allview',all);
 $('#all').textContent=all?'한 장씩 보기':'전체 보기';
 $('#crumb').textContent=all?'전체 교안':id==='home'?'메인 화면':$('#'+id+' h2').textContent;
 $$('[data-chapter]').forEach(a=>{const selected=a.dataset.chapter===id;a.classList.toggle('active',selected);if(selected)a.setAttribute('aria-current','page');else a.removeAttribute('aria-current')});
 $$('.toc-group').forEach(d=>{d.hidden=false;d.open=!!d.querySelector('[aria-current="page"]')});
 $('#rail').classList.remove('open');$('#menu').setAttribute('aria-expanded','false');
 if(view&&requested!==id){let ancestor=target.parentElement;while(ancestor&&ancestor!==view){if(ancestor.tagName==='DETAILS')ancestor.open=true;ancestor=ancestor.parentElement}requestAnimationFrame(()=>target.scrollIntoView({block:'start'}))}
}
window.addEventListener('hashchange',()=>{all=false;route();window.scrollTo({top:0,behavior:"instant"})});$('#all').onclick=()=>{all=!all;route();window.scrollTo({top:0,behavior:"instant"})};$('#menu').onclick=()=>{const o=$('#rail').classList.toggle('open');$('#menu').setAttribute('aria-expanded',String(o))};$('#print').onclick=async()=>{all=true;route();$$('details').forEach(d=>d.open=true);$$('img').forEach(i=>i.loading='eager');await Promise.all($$('img').map(i=>i.decode().catch(()=>{})));window.print()};$('#search').oninput=()=>{const q=$('#search').value.trim().toLowerCase();let n=0;$$('[data-chapter]').forEach(a=>{a.hidden=!!q&&!$('#'+a.dataset.chapter).textContent.toLowerCase().includes(q);if(!a.hidden)n++});$$('.toc-group').forEach(d=>{const match=[...d.querySelectorAll('[data-chapter]')].some(a=>!a.hidden);d.hidden=!!q&&!match;d.open=q?match:!!d.querySelector('[aria-current="page"]')});$('#noresults').hidden=n>0};
$('#toc-expand').onclick=()=>$$('.toc-group').forEach(d=>d.open=true);
$('#toc-collapse').onclick=()=>$$('.toc-group').forEach(d=>d.open=false);
$$('.copy').forEach(b=>b.onclick=async()=>{const t=b.closest('.prompt').querySelector('pre').textContent;try{if(navigator.clipboard&&window.isSecureContext)await navigator.clipboard.writeText(t);else{const a=document.createElement('textarea');a.value=t;document.body.appendChild(a);a.select();const ok=document.execCommand('copy');a.remove();if(!ok)throw Error()}notify('프롬프트를 복사했습니다.')}catch(e){const r=document.createRange();r.selectNodeContents(b.closest('.prompt').querySelector('pre'));const s=getSelection();s.removeAllRanges();s.addRange(r);notify('선택한 문장을 복사해 주세요.')}});$$('.zoom').forEach(b=>b.onclick=()=>{const im=b.querySelector('img');$('#lightbox img').src=im.src;$('#lightbox img').alt=im.alt;$('#lightbox').querySelectorAll('.click-target,.click-arrow').forEach(x=>x.remove());b.querySelectorAll('.click-target,.click-arrow').forEach(x=>$('#lightbox .zoom-frame').appendChild(x.cloneNode(true)));$('#lightbox p').textContent=(b.parentElement.querySelector('figcaption p')||b.parentElement.querySelector('figcaption')).textContent;$('#lightbox').classList.toggle('rotated-preview',b.classList.contains('rotated-photo'));$('#lightbox').showModal()});$('#lightbox .close').onclick=()=>$('#lightbox').close();$('#lightbox').onclick=e=>{if(e.target===$('#lightbox'))$('#lightbox').close()};route();

document.querySelectorAll('a[href="#home"]').forEach(link => link.addEventListener('click', event => {
 event.preventDefault(); all = false;
 document.querySelector('#search').value = '';
 document.querySelectorAll('[data-chapter]').forEach(item => item.hidden = false);
 document.querySelector('#noresults').hidden = true;
 if (location.hash !== '#home') history.pushState(null, '', '#home');
 route(); document.querySelectorAll('video').forEach(video => video.pause());
 window.scrollTo({top: 0, behavior: 'instant'}); document.querySelector('#rail').scrollTop = 0;
}));
window.addEventListener('popstate', () => { all = false; route(); });

const submissionsDialog = document.querySelector('#submissions-dialog');
const submissionsButton = document.querySelector('#open-submissions');
submissionsButton.addEventListener('click', () => submissionsDialog.showModal());
submissionsDialog.querySelector('.submissions-close').addEventListener('click', () => submissionsDialog.close());
submissionsDialog.addEventListener('click', event => {
 if (event.target !== submissionsDialog) return;
 const rect = submissionsDialog.getBoundingClientRect();
 if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) submissionsDialog.close();
});
submissionsDialog.addEventListener('close', () => {
 try { sessionStorage.setItem('workshop-submissions-dismissed', '1'); } catch (_) {}
 submissionsButton.focus({preventScroll:true});
});
let showSubmissions = true;
try { showSubmissions = sessionStorage.getItem('workshop-submissions-dismissed') !== '1'; } catch (_) {}
if (showSubmissions) submissionsDialog.showModal();
