(function(){
const $=(s,r=document)=>r.querySelector(s), $$=(s,r=document)=>[...r.querySelectorAll(s)];
const reduce=matchMedia('(prefers-reduced-motion: reduce)').matches;

/* ---- network canvas */
(function(){
  const c=$('#net'); if(!c||reduce) return;
  const x=c.getContext('2d'); let W,H,P=[],m={x:-999,y:-999};
  function size(){W=c.width=innerWidth;H=c.height=innerHeight;
    const n=Math.min(70,Math.floor(W*H/22000)); P=[...Array(n)].map(()=>({x:Math.random()*W,y:Math.random()*H,vx:(Math.random()-.5)*.35,vy:(Math.random()-.5)*.35}));}
  addEventListener('resize',size); addEventListener('mousemove',e=>{m.x=e.clientX;m.y=e.clientY});
  size();
  (function draw(){
    x.clearRect(0,0,W,H);
    for(const p of P){p.x+=p.vx;p.y+=p.vy;if(p.x<0||p.x>W)p.vx*=-1;if(p.y<0||p.y>H)p.vy*=-1;
      const dx=p.x-m.x,dy=p.y-m.y,d=Math.hypot(dx,dy); if(d<140){p.x+=dx/d*.7;p.y+=dy/d*.7}
      x.beginPath();x.arc(p.x,p.y,1.6,0,6.3);x.fillStyle='rgba(140,170,255,.7)';x.fill();}
    for(let i=0;i<P.length;i++)for(let j=i+1;j<P.length;j++){
      const d=Math.hypot(P[i].x-P[j].x,P[i].y-P[j].y);
      if(d<130){x.strokeStyle='rgba(120,150,255,'+(.22*(1-d/130))+')';x.beginPath();x.moveTo(P[i].x,P[i].y);x.lineTo(P[j].x,P[j].y);x.stroke();}}
    if(!document.hidden) requestAnimationFrame(draw); else setTimeout(()=>requestAnimationFrame(draw),500);
  })();
})();

/* ---- reveal on scroll */
const io=new IntersectionObserver(es=>es.forEach(e=>{if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}}),{threshold:.12});
$$('.rv').forEach(el=>io.observe(el));

/* ---- counters */
const cio=new IntersectionObserver(es=>es.forEach(e=>{
  if(!e.isIntersecting) return; cio.unobserve(e.target);
  const el=e.target,end=+el.dataset.count,suf=el.dataset.suf||'',t0=performance.now(),dur=1600;
  (function f(t){const p=Math.min(1,(t-t0)/dur),v=Math.round(end*(1-Math.pow(1-p,3)));el.textContent=v+suf;if(p<1)requestAnimationFrame(f)})(t0);
}),{threshold:.5});
$$('[data-count]').forEach(el=>cio.observe(el));

/* ---- rotating hero word */
const w=$('#word'); if(w&&!reduce){const words=JSON.parse(w.dataset.words);let i=0;
  setInterval(()=>{i=(i+1)%words.length;w.style.animation='none';void w.offsetWidth;w.textContent=words[i];w.style.animation=''},2600);}

/* ---- card tilt + glow */
$$('.card').forEach(c=>{
  c.addEventListener('mousemove',e=>{const r=c.getBoundingClientRect(),px=(e.clientX-r.left)/r.width,py=(e.clientY-r.top)/r.height;
    c.style.setProperty('--mx',px*100+'%');c.style.setProperty('--my',py*100+'%');
    if(!reduce&&!c.classList.contains('soon')) c.style.transform=`perspective(900px) rotateY(${(px-.5)*9}deg) rotateX(${(.5-py)*9}deg) translateY(-6px)`;});
  c.addEventListener('mouseleave',()=>{c.style.transform=''});
});

/* ---- dropzones */
$$('.drop').forEach(d=>{
  const inp=$('input[type=file]',d),b=$('b',d),s=$('span',d),orig=[b.textContent,s.textContent];
  ['dragenter','dragover'].forEach(v=>d.addEventListener(v,()=>d.classList.add('over')));
  ['dragleave','drop'].forEach(v=>d.addEventListener(v,()=>d.classList.remove('over')));
  inp.addEventListener('change',()=>{const f=inp.files[0];
    if(f){d.classList.add('has');b.textContent=f.name;s.textContent=(f.size/1024).toFixed(f.size>1e6?0:1)+' KB - ready'}
    else{d.classList.remove('has');b.textContent=orig[0];s.textContent=orig[1]}});
});

/* ---- live risk meter (sampling) */
const T=$('#tables');
if(T){
  const t=JSON.parse(T.textContent),form=$('form[data-ajax]');
  const val=n=>{const e=form.querySelector(`[name=${n}]:checked`);return e?e.value:null};
  const circ=2*Math.PI*45;
  function upd(){
    const sig=form.querySelector('[name=sig]').checked,ir=val('ir'),toc=val('toc'),sap=val('sap');
    const rmm=sig?'High':t.rmm[ir][t.toc.indexOf(toc)], sapE=sig?'None':sap, rf=t.rf[rmm][t.sap.indexOf(sapE)];
    const conf=1-Math.exp(-rf);
    $('#m-rmm').textContent=rmm;$('#m-rmm').className='rmm '+rmm;
    $('#m-rf').textContent=rf.toFixed(2);$('#m-ef').textContent=t.ef[rmm].toFixed(1);
    $('#m-sap').textContent=sapE;$('#m-conf').textContent=Math.round(conf*100)+'%';
    $('#m-ring').style.strokeDashoffset=circ*(1-conf);
    const m=val('method');$('#m-method').textContent=m;
    $('#m-note').textContent=sig?'Significant risk: RMM forced to High, no credit for analytics.':(toc==='Tested - effective'?'Controls reliance applied.':'No controls reliance.');
  }
  $('#m-ring').style.strokeDasharray=circ;
  form.addEventListener('change',upd);upd();
}

/* ---- ajax submit with animated overlay */
const fm=$('form[data-ajax]');
if(fm){
  const ov=$('#ov'),box=$('#ovbox'),err=$('#err'),steps=JSON.parse(fm.dataset.steps||'[]');
  const sleep=ms=>new Promise(r=>setTimeout(r,ms));
  function busy(){
    box.innerHTML=`<div class="spin"></div><h3 style="margin:0">Building your workpaper</h3><div class="pbar"><i id="pb"></i></div><div class="stl">${steps.map(s=>`<div>${s}</div>`).join('')}</div>`;
    ov.classList.add('show');const L=$$('.stl div',box);let i=0;
    const tick=()=>{L.forEach((l,k)=>{l.className=k<i?'done':(k===i?'act':'')});$('#pb').style.width=Math.min(94,(i+1)/L.length*100)+'%';i++};
    tick();const h=setInterval(()=>{if(i<L.length)tick()},550);return ()=>{clearInterval(h);L.forEach(l=>l.className='done');$('#pb').style.width='100%'};
  }
  function confetti(){if(reduce)return;const cols=['#6c8cff','#22d3ee','#a78bfa','#34d399','#fbbf24','#fb7185'];
    for(let i=0;i<70;i++){const e=document.createElement('i');e.className='cf';e.style.left=Math.random()*100+'vw';e.style.background=cols[i%6];
      e.style.animationDuration=(1.8+Math.random()*1.8)+'s';e.style.animationDelay=Math.random()*.5+'s';e.style.borderRadius=Math.random()>.5?'50%':'2px';
      document.body.appendChild(e);setTimeout(()=>e.remove(),4200)}}
  const LOCAL=!!(window.AF&&window.AF.local);
  let worker=null,wready=null,wid=0,pending={};
  function engine(){
    if(wready) return wready;
    const st=document.getElementById('engine'); if(st) st.textContent='Loading the calculation engine (one-time, about 15 MB)...';
    worker=new Worker('assets/pyworker.js',{type:'module'});
    wready=new Promise((res,rej)=>{
      worker.onmessage=e=>{const d=e.data;
        if(d.type==='ready'){if(st)st.textContent='Engine ready - runs entirely in your browser';res()}
        else if(d.type==='fatal'){if(st)st.textContent='Engine failed to load';rej(new Error(d.error))}
        else if(d.id&&pending[d.id]){pending[d.id](d);delete pending[d.id]}};
      worker.onerror=e=>rej(new Error(e.message||'worker error'));
      worker.postMessage({cmd:'init'});
    });
    return wready;
  }
  if(LOCAL){const b=fm.querySelector('button.submit');if(b){const d=document.createElement('div');d.id='engine';d.className='hint';d.style.cssText='text-align:center;margin-top:10px';b.after(d)}engine().catch(()=>{});}
  async function localRun(){
    await engine();
    const f=fm.querySelector('input[type=file]').files[0],fields={};
    for(const [k,v] of new FormData(fm).entries()) if(typeof v==='string') fields[k]=v;
    const buf=await f.arrayBuffer(),id=++wid;
    const p=new Promise(res=>{pending[id]=res});
    worker.postMessage({cmd:'run',id,tool:fm.dataset.tool,fields,filename:f.name,buf},[buf]);
    return p;
  }
  fm.addEventListener('submit',async ev=>{
    ev.preventDefault();err.classList.remove('show');
    if(!fm.reportValidity()) return;
    const done=busy();
    try{
      let blob,nm,sum=[];
      if(LOCAL){
        const [r]=await Promise.all([localRun(),sleep(2300)]);
        if(r.error){ov.classList.remove('show');err.textContent=r.error;err.classList.add('show');err.scrollIntoView({behavior:'smooth',block:'center'});return}
        blob=new Blob([r.bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});nm=r.name;sum=r.summary||[];
      }else{
      const [r]=await Promise.all([fetch(fm.action,{method:'POST',body:new FormData(fm),headers:{'X-Requested-With':'fetch'}}),sleep(2300)]);
      if(!r.ok){let m='Something went wrong.';try{m=(await r.json()).error||m}catch(e){}
        ov.classList.remove('show');err.textContent=m;err.classList.add('show');err.scrollIntoView({behavior:'smooth',block:'center'});return}
      blob=await r.blob();const cd=r.headers.get('Content-Disposition')||'';nm=(/filename="?([^";]+)"?/.exec(cd)||[])[1]||'workpaper.xlsx';
      try{sum=JSON.parse(r.headers.get('X-Summary')||'[]')}catch(e){}
      }
      done();await sleep(450);
      const url=URL.createObjectURL(blob);
      box.innerHTML=`<svg class="okic" viewBox="0 0 100 100"><circle cx="50" cy="50" r="46"/><path d="M30 52l14 14 27-30"/></svg>
        <h3 style="margin:0 0 4px;font-size:22px">Workpaper ready</h3><div style="color:var(--mu);font-size:13px">${nm}</div>
        <div class="sum">${sum.map(x=>`<div><b>${x[1]}</b>${x[0]}</div>`).join('')}</div>
        <a class="btn lg" style="width:100%;justify-content:center" href="${url}" download="${nm}" id="dl">Download Excel</a>
        <div style="margin-top:14px"><a href="#" id="cl" style="color:var(--mu);font-size:13px">Close and edit inputs</a></div>`;
      confetti();$('#cl').onclick=e=>{e.preventDefault();ov.classList.remove('show')};
      const a=document.createElement('a');a.href=url;a.download=nm;document.body.appendChild(a);a.click();a.remove();
    }catch(e){ov.classList.remove('show');err.textContent='Network error - is the app still running? '+e;err.classList.add('show')}
  });
}
})();

(function(){
  const $=s=>document.querySelector(s); const btn=$('#fbBtn'),ov=$('#fb'); if(!btn) return;
  let rating=0; const stars=[...document.querySelectorAll('#stars button')];
  const paint=n=>stars.forEach((s,i)=>s.classList.toggle('on',i<n));
  stars.forEach(s=>{s.onclick=()=>{rating=+s.dataset.v;paint(rating)};s.onmouseenter=()=>paint(+s.dataset.v);s.onmouseleave=()=>paint(rating)});
  btn.onclick=()=>ov.classList.add('show'); $('#fbclose').onclick=()=>ov.classList.remove('show');
  ov.addEventListener('click',e=>{if(e.target===ov)ov.classList.remove('show')});
  $('#fbsend').onclick=async()=>{
    const err=$('#fberr'),send=$('#fbsend'); err.classList.remove('show'); send.disabled=true;
    if(window.AF&&window.AF.local){
      const cat=(document.querySelector('[name=fbcat]:checked')||{}).value,msg=$('#fbmsg').value,contact=$('#fbcontact').value;
      if(!msg&&!rating){err.textContent='Please add a rating or a comment.';err.classList.add('show');send.disabled=false;return}
      const body=`Rating: ${rating||'-'}/5\nType: ${cat}\nPage: ${location.pathname}\nFrom: ${contact||'-'}\n\n${msg}`;
      if(window.AF.feedbackUrl){window.open(window.AF.feedbackUrl,'_blank')}
      else{location.href='mailto:'+window.AF.feedbackEmail+'?subject='+encodeURIComponent('Feedback on '+document.title)+'&body='+encodeURIComponent(body)}
      ov.querySelector('.mbox').innerHTML='<div style="text-align:center"><svg class="okic" viewBox="0 0 100 100"><circle cx="50" cy="50" r="46"/><path d="M30 52l14 14 27-30"/></svg><h3 style="margin:0 0 6px;font-size:22px">Almost done</h3><p style="color:var(--mu);margin:0 0 18px">Your email app should open with your feedback ready to send. Press send there.</p><button class="btn" onclick="location.reload()">Close</button></div>';
      return;
    }
    try{
      const r=await fetch('/feedback',{method:'POST',headers:{'Content-Type':'application/json','X-Requested-With':'fetch'},
        body:JSON.stringify({rating,category:(document.querySelector('[name=fbcat]:checked')||{}).value,message:$('#fbmsg').value,contact:$('#fbcontact').value,page:location.pathname})});
      const j=await r.json().catch(()=>({}));
      if(!r.ok){err.textContent=j.error||'Could not send feedback.';err.classList.add('show');send.disabled=false;return}
      ov.querySelector('.mbox').innerHTML='<div style="text-align:center"><svg class="okic" viewBox="0 0 100 100"><circle cx="50" cy="50" r="46"/><path d="M30 52l14 14 27-30"/></svg><h3 style="margin:0 0 6px;font-size:22px">Thank you!</h3><p style="color:var(--mu);margin:0 0 18px">Your feedback has been recorded.</p><button class="btn" onclick="location.reload()">Close</button></div>';
    }catch(e){err.textContent='Network error: '+e;err.classList.add('show');send.disabled=false}
  };
})();
