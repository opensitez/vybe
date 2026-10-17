import assert from 'node:assert/strict';
import {readFileSync,writeFileSync,existsSync} from 'node:fs';
import {resolve,join} from 'node:path';
const dir=resolve(process.argv[2]);
const number=x=>{if(typeof x!=='number')throw TypeError('strict Number');return x;};
const integer=(x,min,max)=>{number(x);if(!Number.isInteger(x)||Object.is(x,-0)||x<min||x>max)throw TypeError('strict integer');return x;};
const core=new Set(['test','cast','concat','equals','compare','length','charCodeAt','codePointAt','fromCharCode','fromCodePoint','substring','intoCharCodeArray','fromCharCodeArray']);
async function load(side){
 const meta=JSON.parse(readFileSync(join(dir,side,'complete.json'))),bytes=readFileSync(join(dir,side,'complete.wasm'));
 const module=await WebAssembly.compile(bytes,{builtins:['js-string'],importedStringConstants:'wasm:string-constants'});
 const stats={objects:0};
 const hosts={
 'wasm:js-number':{fromI32:x=>x|0,fromU32:x=>x>>>0,fromF64:number,toF64:number,toI32:x=>integer(x,-2147483648,2147483647)|0,toU32:x=>integer(x,0,4294967295)|0,test:x=>+(typeof x==='number'),testI32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=-2147483648&&x<=2147483647),testU32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=0&&x<=4294967295)},
 'wasm:js-boolean':{fromI32:x=>!!x,test:x=>+(typeof x==='boolean'),cast:x=>{if(typeof x!=='boolean')throw TypeError('strict Boolean');return +x;}},
 'wasm:js-undefined':{test:x=>+(x===undefined)},
 'wasm:js-symbol':{test:x=>+(typeof x==='symbol'),equals:(a,b)=>{if(typeof a!=='symbol'||typeof b!=='symbol')throw TypeError('strict Symbol');return +(a===b);}},
 'wasm:js-bigint':{test:x=>+(typeof x==='bigint')},
 'wasm:js-string':{fromI32:x=>String(x|0),fromU32:x=>String(x>>>0),fromF64:x=>String(number(x)),fromI64:x=>String(BigInt.asIntN(64,x)),fromU64:x=>String(BigInt.asUintN(64,x))},
 'ecma:map':{size:m=>m instanceof Map?m.size:0},
 'ecma:bigint':{add:(a,b)=>a+b,lt:(a,b)=>+(a<b),ge:(a,b)=>+(a>=b)},
 'ecma:object':{keys:Object.keys,new:()=>{stats.objects++;return {};},get:(o,k)=>o==null?undefined:Reflect.get(Object(o),String(k),o),set:(o,k,v)=>Reflect.set(o,String(k),v,o),is:Object.is},
 'ecma:array':{get:(a,i)=>a[i],splice:(a,start,count,...items)=>Array.prototype.splice.call(a,start,count,...items),push:(a,v)=>Array.prototype.push.call(a,v),join:(a,sep)=>Array.prototype.join.call(a,sep),at:(a,i)=>Array.prototype.at.call(a,i),getValue:(a,k)=>a instanceof Map?a.get(k):(a==null?undefined:Reflect.get(Object(a),String(k),a)),setValue:(a,k,v)=>a instanceof Map?a.set(k,v):Reflect.set(a,String(k),v,a),length:a=>a.length|0},
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
 for(const alias of meta.aliases){const f=meta.exports.find(e=>e.name===`binding-factory:${alias.name}`);assert(f);hosts.env[alias.name].value=instance.exports[f.export](...Array(f.params.length).fill(null));}
 const root=meta.exports[0];instance.exports[root.export](...Array(root.params.length).fill(null));
 const entry=meta.exports.find(e=>e.name==='entry');assert(entry);
 return {run:(...args)=>instance.exports[entry.export](...Array(entry.params.length-args.length).fill(null),...args),stats,bytes:bytes.length};
}

const ws=[await load('slice-positive-before'),await load('slice-positive-after')];
let checks=0,oracleChecks=0;
const corpus=JSON.parse(readFileSync(existsSync(join(dir,'SLICE_POSITIVE_JS_CORPUS.json'))?join(dir,'SLICE_POSITIVE_JS_CORPUS.json'):join(dir,'../JS_CORPUS.json'))).queries;
for(const q of corpus){
 const lo=q.lo===null?0:Math.min(Math.max(q.lo<0?q.lo+q.length:q.lo,0),q.length);
 const hi=q.hi===null?q.length:Math.min(Math.max(q.hi<0?q.hi+q.length:q.hi,0),q.length);
 const expected=Array.from({length:q.length},(_,i)=>i);
 expected.splice(lo,Math.max((hi-lo)|0,0));expected.splice(lo,0,90,91);
 assert.equal(q.error,null);assert.deepEqual(q.result,expected);oracleChecks++;
 for(const w of ws){const input=Array.from({length:q.length},(_,i)=>i);assert.equal(w.run(input,q.lo,q.hi,[90,91]),input);assert.deepEqual(input,expected);checks++;}
}
// Keep the existing remove-then-insert alias behavior; native strict numeric
// conversions can reject mixed bounds, so compare both outcomes and mutations.
const bounds=[null,-100,-1,0,1,3,100,-2147483648,2147483647,1.5,NaN,Infinity,-Infinity,'2','bad',true,undefined];
for(const len of [0,1,5,12])for(const lo of bounds)for(const hi of bounds)for(const alias of [false,true]){
 const outcomes=ws.map(w=>{const input=Array.from({length:len},(_,i)=>i);let error=null;try{w.run(input,lo,hi,alias?input:[90,91]);}catch(e){error=e.constructor.name+':'+e.message;}return {input,error};});
 assert.deepEqual(outcomes[0],outcomes[1],`${len} ${String(lo)} ${String(hi)} ${alias}`);checks++;
}
// User observable bound property access must remain in its original order.
for(const lo of [-2,1])for(const hi of [-1,3]){
 const outcomes=ws.map(w=>{let trace=[];const lower={get Ticks(){trace.push('lo');return lo;}},upper={get Ticks(){trace.push('hi');return hi;}};const input=[0,1,2,3];let error=null;try{w.run(input,lower,upper,[90,91]);}catch(e){error=e.constructor.name+':'+e.message;}return {input,error,trace};});assert.deepEqual(outcomes[0],outcomes[1]);assert(outcomes[0].trace.length>0);checks++;
}
const result={checks,oracleChecks,modules:2,node:process.version,scope:'Complete frozen/candidate modules with genuine js-string builtins; integer-bound independent remove/insert oracle, mixed bound errors/partial mutation, alias behavior, and Ticks getter order. No native timing claim.'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
