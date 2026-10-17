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
function encode(s){let out=String.fromCharCode(s.length+32);for(let i=0;i<s.length;i+=3){const a=s.charCodeAt(i),b=i+1<s.length?s.charCodeAt(i+1):0,c=i+2<s.length?s.charCodeAt(i+2):0;for(const x of [Math.floor(a/4),(a%4)*16+Math.floor(b/16),(b%16)*4+Math.floor(c/64),c%64])out+=String.fromCharCode(x===0?96:x+32);}return out+'\n';}
function decode(s){if(!s.length)return false;const n=(s.charCodeAt(0)-32)&63,need=1+Math.ceil(n/3)*4;if(s.length<need)return false;let out='';for(let i=1;i<need;i+=4){const a=(s.charCodeAt(i)-32)&63,b=(s.charCodeAt(i+1)-32)&63,c=(s.charCodeAt(i+2)-32)&63,d=(s.charCodeAt(i+3)-32)&63;for(const x of [a*4+Math.floor(b/16),(b%16)*16+Math.floor(c/4),(c%4)*64+d])if(out.length<n)out+=String.fromCharCode(x);}return out;}
let checks=0;const rows=[];
for(const decoding of [false,true]){
 const ws=[await load(`uu-pack-${decoding}-false`),await load(`uu-pack-${decoding}-true`)];
 const oracle=decoding?decode:encode;
 const check=s=>{const expected=oracle(s);for(const w of ws){assert.equal(w.run(s),expected);checks++;}};
 for(let cp=0;cp<65536;cp++){const s=String.fromCharCode(cp);
  if(decoding){for(let at=0;at<4;at++){let payload=['A','A','A','A'];payload[at]=s;check('#'+payload.join(''));}const n=(cp-32)&63;check(s+'A'.repeat(Math.ceil(n/3)*4));}
  else{for(const input of [s,'A'+s,s+'B',s+'BC','A'+s+'C','AB'+s])check(input);}
 }
 for(let n=0;n<130;n++){const input=String.fromCharCode(...Array.from({length:n},(_,i)=>(i*37)%256));const coded=encode(input);if(decoding){for(let cut=0;cut<=coded.length;cut++)check(coded.slice(0,cut));}else check(input);}
 for(const n of [255,256,4096,55263,55264,55265,65503,65504,65535,65536]){const input='A'.repeat(n);check(decoding?encode(input):input);}
 for(const x of [null,undefined,false,7,1n,Symbol('x')]){const outcomes=[];for(const w of ws){try{outcomes.push({value:w.run(x)});}catch(e){outcomes.push({error:[e.constructor.name,e.message]});}}assert.deepEqual(outcomes[0],outcomes[1]);assert(outcomes[0].error);checks++;}
 for(const w of ws){let calls=0;const object={[Symbol.toPrimitive](){calls++;throw Error('coercion');}};assert.throws(()=>w.run(object));assert.equal(calls,0);checks++;}
 for(const input of decoding?['#0V%T\n','!80``\n','`']:['Cat','Ω','\0\0\0']){
  const outcomes=[];for(const w of ws){let trace='';Object.defineProperty(Array.prototype,'0',{configurable:true,set(value){trace+=String(value)+'|';Object.defineProperty(this,'0',{value,writable:true,enumerable:true,configurable:true});}});try{outcomes.push({value:w.run(input),trace});}finally{delete Array.prototype[0];}}
  assert.deepEqual(outcomes[0],outcomes[1]);checks++;
 }
 rows.push({decode:decoding,bytes:ws.map(w=>w.bytes)});
}
const result={checks,modules:4,node:process.version,rows,scope:'Complete direct frozen/candidate helpers, genuine js-string builtins. Every UTF16 unit in every input group position, decoder sextet/prefix positions, all short truncation lengths, wide characters/lone surrogates, zero mapping, long-length prefix wrapping, strict errors/no coercion and inherited array setters. Independent scalar arithmetic oracle. No native speed claim.'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
