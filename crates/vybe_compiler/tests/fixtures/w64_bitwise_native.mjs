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
 for(const alias of meta.aliases){const f=meta.exports.find(e=>e.name===`binding-factory:${alias.name}`);assert(f);hosts.env[alias.name].value=instance.exports[f.export](...Array(f.params.length).fill(null));}
 const root=meta.exports[0];instance.exports[root.export](...Array(root.params.length).fill(null));
 const entry=meta.exports.find(e=>e.name==='entry');assert(entry);
 return {run:(...args)=>instance.exports[entry.export](...Array(entry.params.length-args.length).fill(null),...args),stats,bytes:bytes.length};
}


function recipe(kind,name,x){switch(kind){case 0:case 6:return x;case 1:case 5:return Math.trunc(x)|0;case 2:{const v=Math.trunc(x)|0;return v<<3|v>>>29;}case 3:{let n=Math.trunc(x)>>>0,count=0;while(n){n&=n-1;count++;}return count;}case 4:return name==='a'?-7:3;}}

const corpus=JSON.parse(readFileSync(join(dir,'W64_BITWISE_CORPUS.json'))).queries;
const kinds=[[1,1],[1,0],[0,1],[2,3],[4,1],[5,5],[6,0],[0,0],[4,4]];
function expected(op,a,b){const x=BigInt(Math.trunc(a)),y=BigInt(Math.trunc(b));return Number(BigInt.asIntN(64,op===0?x&y:op===1?x|y:x^y));}
let checks=0,trapChecks=0,orderChecks=0,kernelChecks=0,literalChecks=0;
for(let op=0;op<3;op++)for(let test=0;test<9;test++){
 const ws=[await load(`bitwise-compiled-${op}-${test}-before`),await load(`bitwise-compiled-${op}-${test}-after`)];
 for(const q of corpus.filter(q=>q.op===op&&q.case===test)){
  const k=kinds[test],want=expected(op,recipe(k[0],'a',q.a),recipe(k[1],'b',q.b));
  assert.equal(q.expected,want);for(const w of ws){assert.equal(w.run(q.a,q.b),want);checks++;}
 }
 for(const[a,b]of [[NaN,1],[1,Infinity],[9223372036854775808,-1],[null,1],[true,1],['7',1],[1,NaN]]){
  const outcomes=ws.map(w=>{w.stats.trace.length=0;let value,error;try{value=w.run(a,b);}catch(e){error=e.constructor.name+':'+e.message;}return{value,error,trace:[...w.stats.trace]};});
  assert.deepEqual(outcomes[0],outcomes[1],`op=${op} case=${test}`);trapChecks++;orderChecks++;
 }
 if(test===5)for(const w of ws){w.stats.trace.length=0;w.run(7,11);assert.deepEqual(w.stats.trace,['a','b']);orderChecks++;}
}
for(let op=0;op<3;op++)for(let mask=0;mask<4;mask++){
 const ws=[await load(`bitwise-kernel-${op}-${mask}-before`),await load(`bitwise-kernel-${op}-${mask}-after`)];
 for(const a of [-2147483648,-17,-1,0,1,17,2147483647])for(const b of [-2147483648,-17,-1,0,1,17,2147483647])for(const w of ws){assert.equal(w.run(a,b),expected(op,a,b));kernelChecks++;}
}
const literals=[[-7,0],[-2147483648,-1],[2147483647,-1],[-1,4294967295],[4294967295,-1],[0,4294967295]];
for(let op=0;op<3;op++)for(let test=0;test<6;test++)for(const fast of[false,true]){const w=await load(`bitwise-literal-${op}-${test}-${fast}`);assert.equal(w.run(null,null),expected(op,...literals[test]));literalChecks++;}
const result={checks,trapChecks,orderChecks,kernelChecks,literalChecks,modules:114,node:process.version,scope:'Complete actual Compiler and independent frozen/candidate kernels; BigInt bitwise oracle, unsigned literal exclusions, strict traps, observed order; genuine js-string builtins; no native timing claim.'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
