import assert from 'node:assert/strict';
import {readFileSync,writeFileSync} from 'node:fs';
import {resolve,join} from 'node:path';
const dir=resolve(process.argv[2]);
const number=x=>{if(typeof x!=='number')throw TypeError('strict Number');return x;};
const integer=(x,min,max)=>{number(x);if(!Number.isInteger(x)||Object.is(x,-0)||x<min||x>max)throw TypeError('strict integer');return x;};
const core=new Set(['test','cast','concat','equals','compare','length','charCodeAt','codePointAt','fromCharCode','fromCodePoint','substring','intoCharCodeArray','fromCharCodeArray']);
async function load(side){
 const meta=JSON.parse(readFileSync(join(dir,side,'complete.json'))),bytes=readFileSync(join(dir,side,'complete.wasm'));
 const module=await WebAssembly.compile(bytes,{builtins:['js-string'],importedStringConstants:'wasm:string-constants'});
 const stats={objects:0,trace:[]};
 const hosts={
 'wasm:js-number':{fromI32:x=>x|0,fromU32:x=>x>>>0,fromF64:number,toF64:number,toI32:x=>integer(x,-2147483648,2147483647)|0,toU32:x=>integer(x,0,4294967295)|0,test:x=>+(typeof x==='number'),testI32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=-2147483648&&x<=2147483647),testU32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=0&&x<=4294967295)},
 'wasm:js-boolean':{fromI32:x=>!!x,test:x=>+(typeof x==='boolean'),cast:x=>{if(typeof x!=='boolean')throw TypeError('strict Boolean');return +x;}},
 'wasm:js-undefined':{test:x=>+(x===undefined)},
 'wasm:js-symbol':{test:x=>+(typeof x==='symbol'),equals:(a,b)=>{if(typeof a!=='symbol'||typeof b!=='symbol')throw TypeError('strict Symbol');return +(a===b);}},
 'wasm:js-bigint':{test:x=>+(typeof x==='bigint')},
 'wasm:js-string':{fromI32:x=>String(x|0),fromU32:x=>String(x>>>0),fromF64:x=>String(number(x)),fromI64:x=>String(BigInt.asIntN(64,x)),fromU64:x=>String(BigInt.asUintN(64,x))},
 'ecma:object':{new:()=>{stats.objects++;return {};},get:(o,k)=>o==null?undefined:Reflect.get(Object(o),String(k),o),set:(o,k,v)=>Reflect.set(o,String(k),v,o),is:Object.is},
 'ecma:set':{has:(s,v)=>{stats.trace.push('has');return Set.prototype.has.call(s,v);},add:(s,v)=>{stats.trace.push('add');return Set.prototype.add.call(s,v);}},
 'ecma:array':{includes:(a,v)=>Array.prototype.includes.call(a,v),at:(a,i)=>Array.prototype.at.call(a,i),getValue:(a,k)=>a instanceof Map?a.get(k):(a==null?undefined:Reflect.get(Object(a),String(k),a)),setValue:(a,k,v)=>a instanceof Map?a.set(k,v):Reflect.set(a,String(k),v,a),length:a=>a.length|0},
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

let checks=0,traceChecks=0,retainedBaselineErrors=0,originalBugReproductions=0;
const values=[null,undefined,false,true,0,-0,7,NaN,Infinity,'','é😀',{},[]];
for(const kind of [0,1,2]){
 const ws=[await load(`collection-private-${kind}-before`),await load(`collection-private-${kind}-after`)];
 const invoke=(w,input,key,value)=>kind===2?w.run(input,key,value):w.run(input,key);
 if(kind===0){for(const existing of values)for(const key of values){const results=[];for(const w of ws){const input=new Set([existing]);const had=input.has(key);w.stats.trace.length=0;assert.equal(w.run(input,key),!had);assert.deepEqual(w.stats.trace,had?['has']:['has','add']);assert(input.has(key));results.push([...input]);checks++;traceChecks++;}assert.deepEqual(results[0],results[1]);}}
 else for(const map of [false,true])for(const tracked of [0,1,2,3])for(const existing of values)for(const key of ['present','missing']){
  const fallback={marker:'default'},results=[];
  for(const w of ws){const input=map?new Map([['present',existing]]):{present:existing};if(tracked)input.__keys=tracked===1?[]:tracked===2?['present']:['missing','present'];
   const old=map?input.get(key):input[key];const expected=kind===1?((input.__keys??[]).includes(key)||(old!==null&&old!==undefined)):(old===null||old===undefined?fallback:old);
   let result,error=null;try{result=invoke(w,input,key,fallback);}catch(e){error=e.constructor.name+':'+e.message;}
   if(error){throw Error(`unexpected native error kind=${kind} map=${map} tracked=${tracked}: ${error}`);}
   else {assert(Object.is(result,expected),`kind=${kind} map=${map} tracked=${tracked} existing=${String(existing)} key=${key}`);if(kind===2)assert(Object.is(map?input.get(key):input[key],expected));}
   results.push({result,error,values:map?[...input]:Object.entries(input)});checks++;}
  assert.deepEqual(results[0],results[1]);
 }
 if(kind>0)for(const existing of [null,undefined,0,false,{}])for(const tracked of [null,[],['present']]){
 const outcomes=ws.map(w=>{const trace=[];const target={__keys:tracked,present:existing};const input=new Proxy(target,{get(o,k,r){trace.push('get:'+String(k));return Reflect.get(o,k,r);},set(o,k,v,r){trace.push('set:'+String(k));return Reflect.set(o,k,v,r);}});let result,error=null;try{result=invoke(w,input,'present',91);}catch(e){error=e.constructor.name+':'+e.message;}return {result,error,trace,stored:target.present};});assert.deepEqual(outcomes[0],outcomes[1]);assert(outcomes[0].trace.length>0);traceChecks++;
 }
 for(const input of [null,undefined,7,false,'x']){const outcomes=ws.map(w=>{try{return {result:invoke(w,input,'missing',91)};}catch(e){return {error:e.constructor.name+':'+e.message};}});assert.deepEqual(outcomes[0],outcomes[1]);checks++;}
}
const rawHas=await load('collection-private-1-original'),rawDefault=await load('collection-private-2-original');
assert.throws(()=>rawHas.run({__keys:['present'],present:null},'present'),/strict Boolean/);originalBugReproductions++;
for(const input of [{},{present:undefined},new Map(),new Map([['present',undefined]])]){assert.equal(rawDefault.run(input,'present',91),undefined);originalBugReproductions++;}
const result={checks,traceChecks,retainedBaselineErrors,originalBugReproductions,modules:8,node:process.version,scope:'Complete frozen/candidate modules using real native Set/Map/Object adapters and genuine js-string builtins. Return identity, SameValueZero membership, mutations, null/undefined default semantics, invalid receivers, exact has/add order, getter/setter order. Tracked membership predicates execute with the native Boolean/i32 ABI. No retained baseline Boolean-cast errors. No V8 timing claim.'};if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
