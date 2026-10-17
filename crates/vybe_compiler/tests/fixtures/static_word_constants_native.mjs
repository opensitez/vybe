import assert from 'node:assert/strict';
import {readFileSync,writeFileSync,existsSync} from 'node:fs';
import {resolve,join} from 'node:path';
const dir=resolve(process.argv[2]);
let conversionTrace=null;
const number=x=>{if(typeof x!=='number')throw TypeError('strict Number');return x;};
const integer=(x,min,max)=>{number(x);if(!Number.isInteger(x)||Object.is(x,-0)||x<min||x>max)throw TypeError('strict integer');return x;};
const core=new Set(['test','cast','concat','equals','compare','length','charCodeAt','codePointAt','fromCharCode','fromCodePoint','substring','intoCharCodeArray','fromCharCodeArray']);
async function load(side){
 const meta=JSON.parse(readFileSync(join(dir,side,'complete.json'))),bytes=readFileSync(join(dir,side,'complete.wasm'));
 const module=await WebAssembly.compile(bytes,{builtins:['js-string'],importedStringConstants:'wasm:string-constants'});
 const stats={objects:0,trace:[]};
 const hosts={
 'wasm:js-number':{fromI32:x=>x|0,fromU32:x=>x>>>0,fromF64:number,toF64:x=>{if(conversionTrace)conversionTrace.push(typeof x==='number'?Object.is(x,-0)?'-0':String(x):typeof x);return number(x);},toI32:x=>integer(x,-2147483648,2147483647)|0,toU32:x=>integer(x,0,4294967295)|0,test:x=>+(typeof x==='number'),testI32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=-2147483648&&x<=2147483647),testU32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=0&&x<=4294967295)},
 'wasm:js-boolean':{fromI32:x=>!!x,test:x=>+(typeof x==='boolean'),cast:x=>{if(typeof x!=='boolean')throw TypeError('strict Boolean');return +x;}},
 'wasm:js-undefined':{test:x=>+(x===undefined)},
 'wasm:js-symbol':{test:x=>+(typeof x==='symbol'),equals:(a,b)=>{if(typeof a!=='symbol'||typeof b!=='symbol')throw TypeError('strict Symbol');return +(a===b);}},
 'wasm:js-bigint':{test:x=>+(typeof x==='bigint')},
 'wasm:js-string':{fromI32:x=>String(x|0),fromU32:x=>String(x>>>0),fromF64:x=>String(number(x)),fromI64:x=>String(BigInt.asIntN(64,x)),fromU64:x=>String(BigInt.asUintN(64,x))},
 'test:wide-words':{read:(value,label)=>{stats.trace.push(label);return value;}},
 'test:word-fold':{record:value=>{stats.trace.push(value);return null;}},
 'ecma:object':{new:()=>{stats.objects++;return {};},get:(o,k)=>o==null?undefined:Reflect.get(Object(o),String(k),o),set:(o,k,v)=>Reflect.set(o,String(k),v,o),is:Object.is},
 'ecma:array':{at:(a,i)=>Array.prototype.at.call(a,i),getValue:(a,k)=>a instanceof Map?a.get(k):(a==null?undefined:Reflect.get(Object(a),String(k),a)),setValue:(a,k,v)=>a instanceof Map?a.set(k,v):Reflect.set(a,String(k),v,a),length:a=>a.length|0},
 'vybe:js-array':{newWithLength:n=>Array(n)},
 'ecma:function':{apply:(f,receiver,args)=>Reflect.apply(f,receiver,args)},
 'ecma:boolean':{toBoolean:x=>!!x},
 'ecma:string':{length:s=>String(s).length,charCodeAt:(s,i)=>String(s).charCodeAt(i),charAt:(s,i)=>String(s).charAt(i),slice:(s,a,b)=>String(s).slice(a,b),indexOf:(s,t,i)=>String(s).indexOf(t,i),String,fromCharCode:String.fromCharCode,concat:(a,b)=>String.prototype.concat.call(a,b),toUpperCase:s=>String(s).toUpperCase(),padStart:(s,n,p)=>String(s).padStart(n,p),padEnd:(s,n,p)=>String(s).padEnd(n,p)},
 'ecma:number':{Number,toFixed:(n,p)=>Number.prototype.toFixed.call(n,p),toExponential:(n,p)=>Number.prototype.toExponential.call(n,p),toString:(n,r)=>Number.prototype.toString.call(n,r),toPrecision:(n,p)=>Number.prototype.toPrecision.call(n,p)},
 'ecma:math':{trunc:Math.trunc,abs:Math.abs,pow:Math.pow},
 'ecma:regexp':{new:(p,f)=>new RegExp(p,f),replaceAll:(s,re,r)=>s.replaceAll(re,r)},
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
 assert.deepEqual(imports.filter(i=>!Object.hasOwn(hosts[i.module]??{},i.name)),[],'no missing/stubbed imports');
 const instance=await WebAssembly.instantiate(module,hosts);
 for(const alias of meta.aliases){if(!Object.hasOwn(hosts.env,alias.name))continue;const f=meta.exports.find(e=>e.name===`binding-factory:${alias.name}`);assert(f);hosts.env[alias.name].value=instance.exports[f.export](...Array(f.params.length).fill(null));}
 const root=meta.exports[0];instance.exports[root.export](...Array(root.params.length).fill(null));
 const entry=meta.exports.find(e=>e.name==='entry');assert(entry);
 return {run:(...args)=>instance.exports[entry.export](...Array(entry.params.length-args.length).fill(null),...args),stats,bytes:bytes.length};
}




function step(op,a,b){let x=BigInt.asIntN(32,a),y=BigInt.asIntN(32,BigInt(b)),u=BigInt.asUintN(32,x),v=BigInt.asUintN(32,y),k=v&31n,z;
switch(op){case 0:z=x+y;break;case 1:z=x-y;break;case 2:z=x*y;break;case 3:z=x/y;break;case 4:z=u/v;break;case 5:z=x%y;break;case 6:z=u%v;break;case 7:z=x<<k;break;case 8:z=x>>k;break;case 9:z=u>>k;break;case 10:z=x&y;break;case 11:z=x|y;break;case 12:z=x^y;break;case 13:z=+(x===y);break;case 14:z=+(x!==y);break;case 15:z=+(x<y);break;case 16:z=+(u<v);break;case 17:z=+(x<=y);break;case 18:z=+(u<=v);break;case 19:z=+(x>y);break;case 20:z=+(u>v);break;case 21:z=+(x>=y);break;case 22:z=+(u>=v);break;case 23:z=(u<<k)|(u>>(32n-k));break;case 24:z=(u>>k)|(u<<(32n-k));break;case 25:z=[...u.toString(2)].filter(c=>c==='1').length;break;case 26:z=32-u.toString(2).length+(u===0n?1:0);break;case 27:z=u===0n?32:u.toString(2).match(/0*$/)[0].length;break;default:throw Error('op');}return BigInt.asIntN(32,BigInt(z));}


const {rows,inputs}=JSON.parse(readFileSync(join(dir,'STATIC_WORD_GATES.json')));let checks=0;
function effective(k,u){return u?({5:6,15:16,17:18,19:20,21:22}[k]??k):k;}
for(const {unsigned,queries} of rows)for(const fast of [false,true]){
 const w=await load(`static-word-${unsigned}-${fast}`);
 for(const a of inputs){w.stats.trace.length=0;w.run(a);const want=queries.map(([k,b])=>{const n=Number(step(effective(k,unsigned),BigInt.asIntN(32,BigInt(Math.trunc(a))),b));return k>=13?!!n:n;});assert.deepEqual(w.stats.trace,want);checks+=want.length;}
}

const constants=JSON.parse(readFileSync(join(dir,'STATIC_WORD_CONSTANTS.json')));
for(const {unsigned,queries}of constants.rows)for(const fast of[false,true]){const w=await load(`static-word-constant-${unsigned}-${fast}`);w.stats.trace.length=0;w.run(0);const want=queries.map(([k,a,b])=>{const n=Number(step(effective(k,unsigned),BigInt.asIntN(32,BigInt(a)),b));return k>=13?!!n:n;});assert.deepEqual(w.stats.trace,want);checks+=want.length;}
for(let i=0;i<13;i++){const sides=[];for(const fast of[false,true]){const w=await load(`static-word-edge-${i}-${fast}`),results=[];for(const a of[NaN,Infinity,-Infinity,9223372036854775808,-9223372036854775808,-2147483648,-0,null,undefined,true,'3']){w.stats.trace.length=0;try{results.push({value:w.run(a)});}catch(e){results.push({error:e.constructor.name,message:e.message});}assert.deepEqual(w.stats.trace,[77]);checks++;}sides.push(results);}assert.deepEqual(sides[0],sides[1]);}

const carrier=[];for(const fast of[false,true]){const w=await load(`static-word-carrier-${fast}`),results=[];for(const a of[-1,0,0.75,1,4294967295]){try{results.push({value:w.run(a)});}catch(e){results.push({error:e.constructor.name,message:e.message});}checks++;}carrier.push(results);}assert.deepEqual(carrier[0],carrier[1]);
const result={checks,modules:36,node:process.version,scope:'Actual resolved static i32 opcodes, signed/unsigned profiles, independent BigInt oracle, exact boolean carriers, all previous optimizations enabled on both sides; genuine js-string builtins'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
