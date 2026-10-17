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
 const stats={objects:0};
 const hosts={
 'wasm:js-number':{fromI32:x=>x|0,fromU32:x=>x>>>0,fromF64:number,toF64:x=>{if(conversionTrace)conversionTrace.push(typeof x==='number'?Object.is(x,-0)?'-0':String(x):typeof x);return number(x);},toI32:x=>integer(x,-2147483648,2147483647)|0,toU32:x=>integer(x,0,4294967295)|0,test:x=>+(typeof x==='number'),testI32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=-2147483648&&x<=2147483647),testU32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=0&&x<=4294967295)},
 'wasm:js-boolean':{fromI32:x=>!!x,test:x=>+(typeof x==='boolean'),cast:x=>{if(typeof x!=='boolean')throw TypeError('strict Boolean');return +x;}},
 'wasm:js-undefined':{test:x=>+(x===undefined)},
 'wasm:js-symbol':{test:x=>+(typeof x==='symbol'),equals:(a,b)=>{if(typeof a!=='symbol'||typeof b!=='symbol')throw TypeError('strict Symbol');return +(a===b);}},
 'wasm:js-bigint':{test:x=>+(typeof x==='bigint')},
 'wasm:js-string':{fromI32:x=>String(x|0),fromU32:x=>String(x>>>0),fromF64:x=>String(number(x)),fromI64:x=>String(BigInt.asIntN(64,x)),fromU64:x=>String(BigInt.asUintN(64,x))},
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

const corpus=JSON.parse(readFileSync(existsSync(join(dir,'W64_COMPARE_CORPUS.json'))?join(dir,'W64_COMPARE_CORPUS.json'):join(dir,'../CORPUS.json'))).queries;
const cmp=(op,a,b)=>{let x=BigInt.asUintN(64,BigInt(Math.trunc(a))),y=BigInt.asUintN(64,BigInt(Math.trunc(b)));if([2,4,6,8].includes(op)){x=BigInt.asIntN(64,x);y=BigInt.asIntN(64,y);}switch(op){case 0:return x===y;case 1:return x!==y;case 2:case 3:return x<y;case 4:case 5:return x<=y;case 6:case 7:return x>y;case 8:case 9:return x>=y;}};
let checks=0,errorChecks=0,traceChecks=0,compiledChecks=0;
const bad=[NaN,Infinity,-Infinity,-9223372036854777856,18446744073709551616,null,undefined,true,'7',7n,{},Symbol('x'),-7];
for(let op=0;op<10;op++){
 const ws=[await load(`w64-compare-${op}-before`),await load(`w64-compare-${op}-after`)];
 for(const q of corpus.filter(q=>q.op===op)){const expected=cmp(op,q.a,q.b);assert.equal(q.expected,expected);for(const w of ws){assert.equal(+w.run(q.a,q.b),+expected);checks++;}}
 for(const a of bad)for(const b of bad){const outcomes=ws.map(w=>{conversionTrace=[];let result,error=null;try{result=w.run(a,b);}catch(e){error=e.constructor.name+':'+e.message;}const trace=conversionTrace;conversionTrace=null;return {result,error,trace};});assert.deepEqual(outcomes[0],outcomes[1],`op=${op} aType=${typeof a} bType=${typeof b}`);errorChecks++;traceChecks++;}
 for(const[a,b]of [[7,11],[-17,9223372036854775808],[9223372036854775808,18446744073709549568]]){const outcomes=ws.map(w=>{conversionTrace=[];const value=w.run(a,b),trace=conversionTrace;conversionTrace=null;return {value,trace};});assert.deepEqual(outcomes[0],outcomes[1]);traceChecks++;}
 const literals=[[-1.5,0],[-9223372036854775808,1],[9223372036854775808,-1],[18446744073709549568,9223372036854775808],[7,7]];
 for(let n=0;n<literals.length;n++){const w=await load(`w64-compiled-${op}-${n}`);assert.equal(+w.run(),+cmp(op,...literals[n]));compiledChecks++;}
}
const result={checks,errorChecks,traceChecks,compiledChecks,modules:70,node:process.version,scope:'Complete frozen/candidate modules and actual compiled integer expressions; genuine js-string builtins, independent BigInt word/ordering oracle, strict conversion/trap outcomes and exact Number.toF64 conversion order. No native timing claim.'};if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
