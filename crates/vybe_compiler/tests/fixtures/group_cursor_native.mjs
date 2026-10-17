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
let checks=0;const rows=[];const badOutcomes=[];
const outcome=f=>{try{return {value:f()};}catch(e){return {error:[e.constructor.name,e.message]};}};
const oracle=(s,sep,dec)=>{const dot=s.indexOf('.'),end=dot<0?s.length:dot,start=+(s.startsWith('+')||s.startsWith('-'));const digits=s.substring(start,end);let out=s.substring(0,start);for(let i=0;i<digits.length;i++){if(i>0&&(digits.length-i)%3===0)out+=sep;out+=digits.substring(i,i+1);}return out+(dot<0?'':dec+s.substring(dot+1));};
for(const ascii of [false,true])for(const fast of [false,true]){
 const w=await load(`group-cursor-${ascii}-${fast}`);
 const check=(s,sep=',',dec='.')=>{assert.equal(w.run(s,sep,dec),oracle(s,sep,dec),JSON.stringify({s,sep,dec}));checks++;};
 for(let n=0;n<90;n++)for(const sign of ['','+','-'])for(const frac of ['','.','.75'])for(const [sep,dec]of [['',':'],[',','.'],['😀','界']])check(sign+'1234567890'.repeat(9).slice(0,n)+frac,sep,dec);
 if(!ascii){for(let cp=0;cp<65536;cp++)check('12'+String.fromCharCode(cp)+'3456');for(const s of ['😀😀😀','+é😀123.界😀','123😀456','a.b.c','\ud8001234','12345\udc00','a\ud800\udc00b'])check(s);}
 const errors=[];let coercions=0;const object={[Symbol.toPrimitive](){coercions++;return ',';}};
 for(const x of [undefined,null,1,true,Symbol('sep'),object,new String(',')])for(const s of ['','123','1234','1234.5']){errors.push(outcome(()=>w.run(s,x,'.')));errors.push(outcome(()=>w.run(s,',',x)));checks+=2;}
 assert.equal(coercions,0);checks++;badOutcomes.push({ascii,fast,errors});rows.push({ascii,fast,bytes:w.bytes});
}
for(const ascii of [false,true])assert.deepEqual(badOutcomes.find(x=>x.ascii===ascii&&!x.fast).errors,badOutcomes.find(x=>x.ascii===ascii&&x.fast).errors);checks+=2;
const result={checks,modules:4,node:process.version,rows,scope:'Complete generic and ASCII grouping recipes; all BMP UTF16 units and surrogate boundaries in generic path, sign/fraction/group boundaries, Unicode separators, unused invalid separators, exact error class/message, and no object ToPrimitive calls. Genuine js-string builtins. Independent per-unit grouping oracle. No native timing claim.'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
