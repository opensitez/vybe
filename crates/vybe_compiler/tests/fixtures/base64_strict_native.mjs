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
 'ecma:array':{getValue:(a,k)=>a instanceof Map?a.get(k):(a==null?undefined:Reflect.get(Object(a),String(k),a)),setValue:(a,k,v)=>a instanceof Map?a.set(k,v):Reflect.set(a,String(k),v,a),length:a=>a.length|0},
 'vybe:js-array':{newWithLength:n=>Array(n)},
 'ecma:function':{apply:(f,receiver,args)=>Reflect.apply(f,receiver,args)},
 'ecma:boolean':{toBoolean:x=>!!x},
 'ecma:string':{atob:globalThis.atob,indexOf:(s,t,start)=>s.indexOf(t,start)},
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
let checks=0;const rows=[];const badInputs=[];
const outcome=f=>{try{return {value:f()};}catch(e){return {error:[e.constructor.name,e.message]};}};
const oracle=(s,strict)=>strict&&/[^A-Za-z0-9+/=]/.test(s)?false:atob(strict?s:s.replace(/[^A-Za-z0-9+/=]+/g,''));
for(const fast of [false,true]){
 const w=await load(`base64-strict-${fast}`);
 const check=(s,...flags)=>{const strict=flags.length?flags[0]:true;assert.deepEqual(outcome(()=>w.run(s,strict)),outcome(()=>oracle(s,strict)),JSON.stringify(s));checks++;};
 for(let cp=0;cp<65536;cp++){check(`AA${String.fromCharCode(cp)}A`);}
 for(const s of ['', 'AA==','TWFu','aGVsbG8=','YWJj'.repeat(256),'!'+ 'A'.repeat(4095),'A'.repeat(2048)+'!'+'A'.repeat(2047),'A'.repeat(4095)+'!','AA\ud800A','AA😀A','AAAA====','A','AA=A','é'])check(s);
 for(const strict of [false,null,undefined,0,-0,NaN,'',0n])for(const s of ['T W\nFu','aGV!sbG8=',''])check(s,strict);
 for(const strict of [true,1,1n,'yes',Symbol('flag'),{}])for(const s of ['TWFu','!AAA'])check(s,strict);
 let coercions=0;const flag={[Symbol.toPrimitive](){coercions++;throw Error('must not coerce strict flag');}};check('TWFu',flag);assert.equal(coercions,0);checks++;
 const errors=[];for(const input of [undefined,null,1,true,{},Symbol('input')])for(const flag of [true,false]){const r=outcome(()=>w.run(input,flag));assert(r.error);errors.push(r);checks++;}badInputs.push(errors);
 rows.push({fast,bytes:w.bytes});
}
assert.deepEqual(badInputs[0],badInputs[1]);checks++;
const result={checks,modules:2,node:process.version,rows,scope:'Actual full two-argument strict/lenient decoder recipes in complete modules. All65536 UTF16 units in payloads, lone surrogates/astral text, valid and malformed padding, first/middle/last invalid units, runtime strict flags, unchanged nonstring errors and no flag ToPrimitive. Genuine js-string builtins and real atob/RegExp host bridges. Independent membership regex plus atob oracle. No native speed claim.'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
