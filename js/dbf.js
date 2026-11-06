// dbf.js
// Membaca file .dbf (tabel atribut shapefile) menjadi array objek: {cols, rows}.

function parseDBF(buf){
 const dv=new DataView(buf),n=dv.getUint32(4,true),hl=dv.getUint16(8,true),rl=dv.getUint16(10,true);
 const dec=new TextDecoder('utf-8'),u=new Uint8Array(buf),fs=[];
 for(let o=32;u[o]!==0x0D;o+=32){let nm='';for(let i=0;i<11&&u[o+i];i++)nm+=String.fromCharCode(u[o+i]);fs.push({nm,l:u[o+16]})}
 const rows=[];
 for(let i=0;i<n;i++){const b=hl+i*rl;if(u[b]===0x2A)continue;let p=b+1;const r={};
  for(const f of fs){r[f.nm]=dec.decode(u.subarray(p,p+f.l)).replace(/\0/g,'').trim();p+=f.l}rows.push(r)}
 return{cols:fs.map(f=>f.nm),rows};
}
