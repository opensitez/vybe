import assert from 'node:assert/strict';
import {readFileSync,writeFileSync} from 'node:fs';
import {resolve,join} from 'node:path';
const dir=resolve(process.argv[2]);
const number=x=>{if(typeof x!=='number')throw TypeError('strict Number');return x;};
const integer=(x,min,max)=>{number(x);if(!Number.isInteger(x)||Object.is(x,-0)||x<min||x>max)throw TypeError('strict integer');return x;};
const core=new Set(['test','cast','concat','equals','compare','length','charCodeAt','codePointAt','fromCharCode','fromCodePoint','substring','intoCharCodeArray','fromCharCodeArray']);
async function load(side){
 const meta=JSON.parse(readFileSync(join(dir,side,'complete.json'))),bytes=readFileSync(join(dir,side,'complete.wasm'));
 const module=await WebAssembly.compile(bytes,{builtins:['js-string','text-encoder'],importedStringConstants:'wasm:string-constants'});
 const stats={objects:0};
 const hosts={
 'wasm:js-number':{fromI32:x=>x|0,fromU32:x=>x>>>0,fromF64:number,toF64:number,toI32:x=>integer(x,-2147483648,2147483647)|0,toU32:x=>integer(x,0,4294967295)|0,test:x=>+(typeof x==='number'),testI32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=-2147483648&&x<=2147483647),testU32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=0&&x<=4294967295)},
 'wasm:js-boolean':{fromI32:x=>!!x,test:x=>+(typeof x==='boolean'),cast:x=>{if(typeof x!=='boolean')throw TypeError('strict Boolean');return +x;}},
 'wasm:js-undefined':{test:x=>+(x===undefined)},
 'wasm:js-symbol':{test:x=>+(typeof x==='symbol'),equals:(a,b)=>{if(typeof a!=='symbol'||typeof b!=='symbol')throw TypeError('strict Symbol');return +(a===b);}},
 'wasm:js-bigint':{test:x=>+(typeof x==='bigint')},
 'wasm:js-string':{fromI32:x=>String(x|0),fromU32:x=>String(x>>>0),fromF64:x=>String(number(x)),fromI64:x=>String(BigInt.asIntN(64,x)),fromU64:x=>String(BigInt.asUintN(64,x))},
 'ecma:object':{new:()=>{stats.objects++;return {};},get:(o,k)=>o==null?undefined:Reflect.get(Object(o),String(k),o),set:(o,k,v)=>Reflect.set(o,String(k),v,o),is:Object.is},
 'ecma:array':{getValue:(a,k)=>a instanceof Map?a.get(k):(a==null?undefined:Reflect.get(Object(a),String(k),a)),setValue:(a,k,v)=>a instanceof Map?a.set(k,v):Reflect.set(a,String(k),v,a),length:a=>a.length|0},
 'vybe:js-array':{newWithLength:n=>Array(n)},
 'ecma:function':{apply:(f,receiver,args)=>Reflect.apply(f,receiver,args)},
 'ecma:boolean':{toBoolean:x=>!!x},
 env:{}
 };
 for(const {module:m,name,detail:d}of meta.inventory)if(d.kind==='global'&&m!=='wasm:string-constants'){
  let value;if(m==='wasm:js-undefined'&&name==='value')value=undefined;
  else if(m==='wasm:js-boolean'&&['true','false'].includes(name))value=name==='true';
  else if(m==='env'&&meta.aliases.some(a=>a.name===name))value=undefined;
  else throw Error(`unresolved global ${m}:${name}`);
  (hosts[m]??={})[name]=new WebAssembly.Global({value:'externref',mutable:d.mutable},value);
 }
 const imports=WebAssembly.Module.imports(module);
 assert.equal(imports.filter(i=>i.module==='wasm:js-string'&&core.has(i.name)).length,0,'genuine js-string builtins');
 assert.equal(imports.filter(i=>i.module==='wasm:text-encoder').length,0,'genuine UTF8 builtins');
 assert.deepEqual(imports.filter(i=>!Object.hasOwn(hosts[i.module]??{},i.name)),[],'no missing/stubbed imports');
 const instance=await WebAssembly.instantiate(module,hosts);
 for(const alias of meta.aliases){const f=meta.exports.find(e=>e.name===`binding-factory:${alias.name}`);assert(f);hosts.env[alias.name].value=instance.exports[f.export](...Array(f.params.length).fill(null));}
 const root=meta.exports[0];instance.exports[root.export](...Array(root.params.length).fill(null));
 const entry=meta.exports.find(e=>e.name==='entry');assert(entry);
 return {run:x=>instance.exports[entry.export](...Array(entry.params.length-1).fill(null),x),stats,bytes:bytes.length};
}
let checks=0; const rows=[]; const errors=[];
for(const fast of [false,true,'fallback']) {
 const w=await load(`byte-length-${fast}`);
 const encoder=new TextEncoder();
 const check=s=>{assert.equal(w.run(s),encoder.encode(s).length|0);checks++;};
 for(let u=0;u<65536;u++)check(String.fromCharCode(u));
 for(let base=65536;base<0x110000;base+=1024){let s='';for(let cp=base;cp<base+1024;cp++)s+=String.fromCodePoint(cp);check(s);}
 for(const s of ['','abc','a\0b','aé界😀','\ud800a\udc00','\ud800\ud800','\udc00\ud800','😀😀','\udbff\udfff'])check(s);
 const sideErrors=[]; for(const x of [undefined,null,true,1,{},Symbol('x')]){let thrown;try{w.run(x);}catch(e){thrown=e;}assert(thrown);sideErrors.push([thrown.constructor.name,thrown.message]);checks++;} errors.push(sideErrors);
 rows.push({fast,bytes:w.bytes});
}
assert.deepEqual(errors[0],errors[1]);assert.deepEqual(errors[0],errors[2]);checks+=2;
const wrapping=await load('byte-length-wrapping');
const unsigned=await load('byte-length-u32-wrapping');
for(const s of ['a','é','界','😀','aé界😀','\0','é😀é😀']) {
 assert.equal(wrapping.run(s),(2147483647+new TextEncoder().encode(s).length)|0);checks++;
 assert.equal(unsigned.run(s),(-1+new TextEncoder().encode(s).length)|0);checks++;
}
const entries=[await load("byte-length-false"),await load("byte-length-true")];
for(const side of [false,true]) {
 const declared=await load(`byte-length-declared-${side}`);
 for(const s of ['','a','é','界','😀','\ud800','aé界😀']) {assert.equal(declared.run(s),new TextEncoder().encode(s).length);checks++;}
 for(const x of [null,undefined,true,1,{}]) {assert.throws(()=>declared.run(x));checks++;}
}
const timings=[];
for(const [label,base] of [["empty",""],["short","abc"],["ascii","a".repeat(256)],["mixed","aé界😀".repeat(64)]]) {
 const inputs=Array.from({length:16},(_,i)=>base?base+String.fromCharCode(65+i):base);
 const samples=[[],[]];let checksum=0;
 for(let pair=0;pair<7;pair++)for(const side of pair%2?[1,0]:[0,1]) {
  const run=entries[side].run;let calls=0;const start=process.hrtime.bigint();const cpu=process.threadCpuUsage();
  do {for(let j=0;j<256;j++){checksum^=run(inputs[calls&15]);calls++;}}while(process.hrtime.bigint()-start<30000000n);
  const used=process.threadCpuUsage(cpu);samples[side].push((used.user+used.system)*1000/calls);
 }
 const median=a=>a.toSorted((x,y)=>x-y)[3];
 timings.push({label,oldCpuNs:median(samples[0]),newCpuNs:median(samples[1]),cpuRatio:median(samples[0])/median(samples[1]),checksum,samples});
}
const result={timings,checks,modules:7,node:process.version,rows,scope:'Complete old/production/forced-fallback/signed-and-unsigned-wrapping-counter/declared-value-ABI Wasm with genuine js-string builtins and stringref instructions; all 65536 individual UTF16 units including lone surrogates, all astral scalar values in 1024-codepoint batches, mixed/surrogate-boundary strings and unchanged nonstring error class/message. TextEncoder is the independent width oracle. Seven alternating thread-CPU pairs also measure changing-input complete native modules.'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
