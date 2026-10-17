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
const ranges=JSON.parse(readFileSync(join(dir,'cased-oracle.json')));
function contains(rs,x){let lo=0,hi=rs.length;while(lo<hi){const m=(lo+hi)>>>1;if(rs[m][1]<=x)lo=m+1;else hi=m;}return lo<rs.length&&rs[lo][0]<=x;}
let checks=0;const rows=[];
for(let cls=0;cls<6;cls++){
 const ws=[await load(`cased-short-${cls}-false`),await load(`cased-short-${cls}-true`)];
 const opposite=cls===4?5:4;
 const check=(s,expected)=>{for(const w of ws){assert.equal(w.run(s),expected,`class=${cls} input=${JSON.stringify(s)}`);checks++;}};
 for(let cp=0;cp<=0x10ffff;cp++){
  if(cls<4&&cp>0xffff&&cp%257!==0)continue;
  const s=String.fromCodePoint(cp),same=contains(ranges[cls],cp),bad=cls>=4&&contains(ranges[opposite],cp);
  check(s,same&&!bad);
  if(cls>=4)check((cls===4?'A':'a')+s,!bad);
 }
 for(const s of ['','ABC','abc','A-a','123','😀','中','A\u0301','ΩΩΩ','ΩAΩ','ΩΜΕΓΑΚΟΣΜΟΣ','ωμέγακόσμος','ÉéÉ','ééé','ßß','𐐀𐐀','𐐨𐐨','Ω 😀','\ud800','\udc00']){
  const cps=Array.from(s,c=>c.codePointAt(0));const expected=cls<4?cps.length>0&&cps.every(x=>contains(ranges[cls],x)):cps.some(x=>contains(ranges[cls],x))&&!cps.some(x=>contains(ranges[opposite],x));
  check(s,expected);
 }
 for(const x of [null,undefined,true,0,1n,Symbol('x')]){const outcomes=[];for(const w of ws){try{outcomes.push({value:w.run(x)});}catch(e){outcomes.push({error:[e.constructor.name,e.message]});}}assert.deepEqual(outcomes[0],outcomes[1]);assert(outcomes[0].error);checks++;}
 for(const w of ws){let calls=0;const obj={[Symbol.toPrimitive](){calls++;throw Error('unexpected coercion');}};assert.throws(()=>w.run(obj));assert.equal(calls,0);checks++;}
 rows.push({class:cls,bytes:ws.map(w=>w.bytes)});
}
const result={checks,modules:12,node:process.version,rows,scope:'Both frozen and candidate complete native modules. Every Unicode codepoint (including lone surrogate code units) for upper/lower, both before and after a matching cased prefix; all BMP plus sampled astral codepoints for four other classifiers. Independent Rust char-property ranges, mixed/empty strings and strict nonstring/no-coercion errors. No native timing claim.'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
