import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
const dir=process.argv[2];
const number=x=>{if(typeof x!=='number')throw TypeError('Number');return x;};
const i32=x=>{number(x);if(!Number.isInteger(x)||Object.is(x,-0)||x< -2147483648||x>2147483647)throw TypeError('i32');return x|0;};
const base={
 'ecma:array':{getValue:(a,k)=>a instanceof Map?a.get(k):(a==null?undefined:Reflect.get(Object(a),String(k),a)),setValue:(a,k,v)=>a instanceof Map?a.set(k,v):Reflect.set(a,String(k),v,a),length:a=>a.length|0},
 'ecma:function':{apply:(f,receiver,args)=>Reflect.apply(f,receiver,args)},
 'ecma:object':{new:()=>({}),get:(o,k)=>o==null?undefined:Reflect.get(Object(o),String(k),o),set:(o,k,v)=>Reflect.set(o,String(k),v,o),is:Object.is},
 'wasm:js-string':{fromI32:x=>String(x|0),fromU32:x=>String(x>>>0),fromF64:x=>String(number(x)),fromI64:x=>String(BigInt.asIntN(64,x)),fromU64:x=>String(BigInt.asUintN(64,x))},
 'wasm:js-number':{fromI32:x=>x|0,fromU32:x=>x>>>0,fromF64:number,toF64:number,toI32:i32,toU32:x=>{number(x);if(!Number.isInteger(x)||Object.is(x,-0)||x<0||x>4294967295)throw TypeError('u32');return x>>>0;},testU32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=0&&x<=4294967295),test:x=>+(typeof x==='number'),testI32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=-2147483648&&x<=2147483647)},
 'wasm:js-boolean':{fromI32:x=>!!x,test:x=>+(typeof x==='boolean'),cast:x=>{if(typeof x!=='boolean')throw TypeError('Boolean');return +x;}},
 'wasm:js-undefined':{test:x=>+(x===undefined)},
 'wasm:js-bigint':{test:x=>+(typeof x==='bigint')},
 'wasm:js-symbol':{test:x=>+(typeof x==='symbol'),equals:(a,b)=>{if(typeof a!=='symbol'||typeof b!=='symbol')throw TypeError('Symbol');return +(a===b);}},
};
// Add a test export for chunk 1. _start is a void startup wrapper.
// Preserve every original instruction and section except this export alias.
const leb=n=>{const a=[];do{let b=n&127;n=Math.floor(n/128);if(n)b|=128;a.push(b);}while(n);return a;};
function directEntry(bytes){
 let p=8;const sections=[];let imports=0;
 const read=()=>{let n=0,shift=0,b;do{b=bytes[p++];n+=(b&127)*2**shift;shift+=7;}while(b&128);return n;};
 const skipName=()=>{const n=read();p+=n;};
 while(p<bytes.length){const id=bytes[p++],length=read(),start=p,end=p+length;
  if(id===2){const count=read();for(let i=0;i<count;i++){skipName();skipName();const kind=bytes[p++];if(kind===0){read();imports++;}else{assert.equal(kind,3);break;}}}
  if(id===7){const count=read(),rest=bytes.subarray(p,end),name=Buffer.from('math_result');const payload=Buffer.from([...leb(count+1),...rest,...leb(name.length),...name,0,...leb(imports+1)]);sections.push(Buffer.from([id,...leb(payload.length),...payload]));}
  else sections.push(Buffer.from([id,...leb(length),...bytes.subarray(start,end)]));
  p=end;
 }
 return Buffer.concat([bytes.subarray(0,8),...sections]);
}
async function load(kind,fast){
 const bytes=fs.readFileSync(path.join(dir,`${kind}-${fast}.wasm`));
 const module=await WebAssembly.compile(directEntry(bytes),{builtins:['js-string'],importedStringConstants:'wasm:string-constants'});
 const hosts=Object.fromEntries(Object.entries(base).map(([k,v])=>[k,{...v}]));
 for(const i of WebAssembly.Module.imports(module)){
  if(i.module==='wasm:string-constants')continue;
  if(i.kind==='global'&&i.module==='wasm:js-undefined'&&i.name==='value')hosts[i.module].value=new WebAssembly.Global({value:'externref',mutable:false},undefined);
  else if(i.kind==='global'&&i.module==='wasm:js-boolean'&&['true','false'].includes(i.name))hosts[i.module][i.name]=new WebAssembly.Global({value:'externref',mutable:false},i.name==='true');
  else if(!(i.name in(hosts[i.module]??{})))throw Error(`Missing actual import ${JSON.stringify(i)}`);
 }
 const instance=await WebAssembly.instantiate(module,hosts);const f=instance.exports.math_result;

 return (...args)=>f(...Array(f.length-args.length).fill(null),...args);
}
const d=new DataView(new ArrayBuffer(8));
const fromBits=s=>{d.setBigUint64(0,BigInt('0x'+s));return d.getFloat64(0);};
const bits=n=>{d.setFloat64(0,n);return d.getBigUint64(0);};
const q=JSON.parse(fs.readFileSync(path.join(dir,'queries.json'),'utf8'));
let checks=0,oracle_checks=0,rejection_checks=0;
for(let kind=0;kind<4;kind++)for(const lane of [0,1]){
 const ws=[await load(`${kind}-${lane}`,false),await load(`${kind}-${lane}`,true)];
 for(const word of q){const x=fromBits(word),a=ws[0](x,2),b=ws[1](x,2);assert.equal(bits(a),bits(b),`${kind} ${lane} ${word}`);checks++;
  if(lane===1&&kind>=1&&kind<=2&&!Number.isNaN(x)){
   const up=kind===1; let expected;
   if(x===(up?Infinity:-Infinity))expected=x;
   else if(x===0)expected=up?Number.MIN_VALUE:-Number.MIN_VALUE;
   else {const step=up===(x>0)?1n:-1n;d.setBigUint64(0,BigInt.asUintN(64,bits(x)+step));expected=d.getFloat64(0);}
   assert.equal(bits(b),bits(expected),`${kind} oracle ${word}`);oracle_checks++;
  }
 }
 for(const bad of [null,undefined,true,'1',{},[],Symbol('x')])for(const args of[[bad,1],[1,bad]]){
  const results=ws.map(f=>{try{return {bits:String(bits(f(...args)))};}catch(e){return {error:e.constructor.name,message:e.message};}});assert.deepEqual(results[0],results[1]);rejection_checks+=2;
 }
}
const result={node:process.version,checks,oracle_checks,rejection_checks,modules:16,passed:true,scope:'Actual complete frozen/candidate nextafter, nextUp, nextDown and ULP modules; exact bits in both lanes, independent ordered F64 neighbor oracle and strict import errors'};
if(process.argv[3])fs.writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');
console.log(JSON.stringify(result));
