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
 'ecma:array':{push:(a,v)=>Array.prototype.push.call(a,v),join:(a,sep)=>Array.prototype.join.call(a,sep),at:(a,i)=>Array.prototype.at.call(a,i),getValue:(a,k)=>a instanceof Map?a.get(k):(a==null?undefined:Reflect.get(Object(a),String(k),a)),setValue:(a,k,v)=>a instanceof Map?a.set(k,v):Reflect.set(a,String(k),v,a),length:a=>a.length|0},
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
function digit(n){return n>=48&&n<=57?n-48:n>=65&&n<=70?n-55:n>=97&&n<=102?n-87:-1;}
function oracle(s){let out='';for(let i=0;i<s.length;){if(s.charCodeAt(i)===61){let j=i+1;while(j<s.length&&[32,9].includes(s.charCodeAt(j)))j++;if(j===s.length)break;const c=s.charCodeAt(j);if(c===13||c===10){j++;if(c===13&&j<s.length&&s.charCodeAt(j)===10)j++;i=j;continue;}if(j===i+1&&i+2<s.length){const a=digit(s.charCodeAt(i+1)),b=digit(s.charCodeAt(i+2));if(a>=0&&b>=0){out+=String.fromCharCode(a*16+b);i+=3;continue;}}}out+=s.charAt(i++);}return out;}
const ws=[await load('qp-cursor-false'),await load('qp-cursor-true')];let checks=0;
const check=s=>{const expected=oracle(s);for(const w of ws){assert.equal(w.run(s),expected);checks++;}};
for(let base=0;base<65536;base+=256){let s='';for(let u=base;u<base+256;u++)s+='='+String.fromCharCode(u)+'0|=0'+String.fromCharCode(u)+'|';check(s);}
const alphabet=['=','0','A','f','G','\r','\n',' ','\t','é'];for(let n=0;n<5;n++)for(let bits=0;bits<10**n;bits++){let k=bits,s='';for(let i=0;i<n;i++){s+=alphabet[k%10];k=Math.floor(k/10);}check(s);}
for(const a of '0123456789ABCDEFabcdef')for(const b of '0123456789ABCDEFabcdef')check('='+a+b);
for(const s of ['', '=', '= ', '=\t\r\nx','abc=\rnext','==41','=00=FF=7f','é😀=41😀é','x\0=00\r\n','=0','=GG=G0=0G'])check(s);
for(const x of [null,undefined,7,false,Symbol('x'),{},[]]){for(const w of ws){assert.throws(()=>w.run(x));checks++;}}
for(const s of ['abc','=aF','= \t\r\n!','=G0','😀']){const outcomes=[];for(const w of ws){let trace='';Object.defineProperty(Array.prototype,'0',{configurable:true,set(value){trace+=String(value)+'|';Object.defineProperty(this,'0',{value,writable:true,enumerable:true,configurable:true});}});try{outcomes.push({value:w.run(s),trace});}finally{delete Array.prototype[0];}}assert.deepEqual(outcomes[0],outcomes[1]);checks++;}
const result={checks,modules:2,node:process.version,scope:'Complete frozen/candidate modules, independent UTF16 scalar reference, every UTF16 unit in both digit positions, malformed/padding/breaks, strict type errors and inherited array setters. Genuine js-string builtins. No native timing claim.'};if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
