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
 const stats={objects:0};
 const hosts={
 'wasm:js-number':{fromI32:x=>x|0,fromU32:x=>x>>>0,fromF64:number,toF64:number,toI32:x=>integer(x,-2147483648,2147483647)|0,toU32:x=>integer(x,0,4294967295)|0,test:x=>+(typeof x==='number'),testI32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=-2147483648&&x<=2147483647),testU32:x=>+(typeof x==='number'&&Number.isInteger(x)&&!Object.is(x,-0)&&x>=0&&x<=4294967295)},
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
let checks=0;const rows=[];const workers=[];
const outcome=f=>{try{return {value:f()};}catch(e){return {error:[e.constructor.name,e.message]};}};
for(const fast of [false,true])workers.push(await load(`sprintf-dispatch-${fast}`));
const check=(fmt,args)=>{assert.deepEqual(outcome(()=>workers[1].run(fmt,args)),outcome(()=>workers[0].run(fmt,args)),JSON.stringify({fmt,args}));checks++;};
for(const conv of 'sdiufFeEgGxXobcBOq')for(const flag of ['','#','+',' ','0','-','#0','+ ',' +'])for(const width of ['','4','12'])for(const precision of ['','.0','.3'])for(const num of [0,-0,.5,-1.5,15.75,255,-255,4294967295,NaN,Infinity,-Infinity])check(`a%${flag}${width}${precision}${conv}z`,[num]);
for(let cp=0;cp<65536;cp++)check('%#'+String.fromCharCode(cp),[255]);
for(const [fmt,args,expected]of [['%+ d',[15],'+15'],['% +d',[15],'+15'],['%+.0d',[0],'+'],['% .0d',[0],' '],['%+ d',[-15],'-15'],['%#x',[255],'0xff'],['%#X',[255],'0XFF'],['%#o',[255],'0377'],['%#b',[255],'11111111'],['%#x',[.5],'0x0'],['%#X',[-15],'0X-F'],['%08x',[255],'000000ff']]){for(const w of workers){assert.equal(w.run(fmt,args),expected);checks++;}}
for(const fmt of ['','plain😀text','%%','%','%#','%s|%d|%X|%o|%b|%g','%2$x/%1$o/%2$X','%08.3X/%#b/%#o'])check(fmt,[15,255,-1,0,1.25,3]);
for(const fmt of ['%12.3s','%#x','%#X','%#o','%#b','%12.3f','%g','%c','%s|%#x|%c|%g']){
 const results=[];for(const w of workers){let trace=[];const obj={[Symbol.toPrimitive](hint){trace.push(hint);return hint==='string'?'word':15;}};const args=new Proxy([obj,255,obj,obj],{get(t,k,r){if(k==='length'||/^\d+$/.test(String(k)))trace.push('get:'+String(k));return Reflect.get(t,k,r);}});results.push({result:outcome(()=>w.run(fmt,args)),trace});}assert.deepEqual(results[0],results[1],fmt);checks++;
}
for(const fmt of ['%s','%4s','%x','%4c'])check(fmt,[Symbol('arg')]);
for(let producer=0;producer<5;producer++){const w=await load(`br-table-selector-${producer}`);for(const x of [-2,-1,0,1,2,2147483647]){const i=producer===0?1:producer===2?(x+1)|0:producer===4?+(x<0):x;assert.equal(w.run(x),i===0?10:i===1?11:99);checks++;}}
for(let i=0;i<workers.length;i++)rows.push({fast:!!i,bytes:workers[i].bytes});
const result={checks,modules:7,node:process.version,rows,scope:'Complete legacy/candidate formatter modules, flag/width/precision/value cross-product, all UTF16 conversion codes, representative independent radix outputs, positional/mixed formats, nonfinite values, observable proxy argument reads and Number/String ToPrimitive order, Symbol errors. Five standard br_table selector producer modules executed. No native speed claim.'};if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
