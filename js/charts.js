// charts.js
// Membuat grafik (bar untuk 5 terbesar, pie untuk penguasaan tanah) memakai Chart.js.

const valPlugin={id:'vals',afterDatasetsDraw(ch){
 const c=ch.ctx,ds=ch.data.datasets[0],meta=ch.getDatasetMeta(0);c.save();c.font='11px Arial';c.textAlign='center';
 meta.data.forEach((el,i)=>{
  if(ch.config.type==='bar'){c.fillStyle='#222';c.fillText(fmt(ds.data[i]),el.x,el.y-5)}
  else{const tot=ds.data.reduce((a,b)=>a+b,0),pc=ds.data[i]/tot*100;if(pc<3)return;const p=el.tooltipPosition();c.fillStyle='#fff';c.font='bold 12px Arial';c.fillText(pc.toFixed(1)+'%',p.x,p.y+4)}
 });c.restore()}};
const bgPlugin={id:'bg',beforeDraw(ch){const c=ch.ctx;c.save();c.fillStyle='#fff';c.fillRect(0,0,ch.width,ch.height);c.restore()}};
function mkChart(cv,def,d,anim=true){
 const pie=def.type==='pie',src=pie?d.body:d.top;
 return new Chart(cv,{type:pie?'pie':'bar',data:{labels:src.map(x=>x.c),datasets:[{data:src.map(x=>x.t),backgroundColor:pie?PAL:NAVY}]},
  options:{responsive:false,animation:anim?undefined:false,layout:{padding:{top:18}},
     plugins: { legend: { display: pie, position: 'right', labels: { boxWidth: 10, font: { size: 10 } } }, title: { display: true, text: yr(def.ct), font:{size:12}}},
   scales:pie?{}:{y:{display:false,beginAtZero:true},x:{ticks:{font:{size:10}}}}},plugins:[bgPlugin,valPlugin]});
}
