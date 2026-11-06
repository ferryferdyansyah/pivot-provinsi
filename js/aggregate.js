// aggregate.js
// Menghitung luas (LUASHA) per kategori x kabupaten/kota, total, persentase, dan 5 terbesar.
// Mengembalikan null jika kolom tidak ada / kosong -> nanti tampil '(Data tidak ditemukan)'.

function agg(f){
 if(!ROWS.length||!COLS.includes(f))return null;
 const kk=r=>(COLS.includes('WADMKC')&&r.WADMKC)||'-';
 const kecs=[...new Set(ROWS.map(kk))].sort(),m={};let any=false;
 for(const r of ROWS){const c=r[f];if(!c)continue;any=true;const k=kk(r);(m[c]??={});m[c][k]=(m[c][k]||0)+(parseFloat(r.LUASHA)||0)}
 if(!any)return null;
 const body=Object.keys(m).sort((a,b)=>a.localeCompare(b)).map(c=>{const v=kecs.map(k=>m[c][k]||0);return{c,v,t:v.reduce((a,b)=>a+b,0)}});
 const T=body.reduce((a,b)=>a+b.t,0);body.forEach(b=>b.p=T?b.t/T*100:0);
 return{kecs,body,T,kt:kecs.map((_,i)=>body.reduce((a,b)=>a+b.v[i],0)),top:[...body].sort((a,b)=>b.t-a.t).slice(0,5)};
}
