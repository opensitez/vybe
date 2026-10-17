import assert from 'node:assert/strict';
import {readFileSync,writeFileSync} from 'node:fs';
import {resolve,join} from 'node:path';
const dir=resolve(process.argv[2]);
const number=x=>{if(typeof x!=='number')throw TypeError('strict Number');return x;};
const integer=(x,min,max)=>{number(x);if(!Number.isInteger(x)||Object.is(x,-0)||x<min||x>max)throw TypeError('strict integer');return x;};
const core=new Set(['test','cast','concat','equals','compare','length','charCodeAt','codePointAt','fromCharCode','fromCodePoint','substring','intoCharCodeArray','fromCharCodeArray']);
async function load(side){
 const meta=JSON.parse(readFileSync(join(dir,side,'complete.json'))),bytes=readFileSync(join(dir,side,'complete.wasm'));
 const module=await WebAssembly.compile(bytes,{builtins:['js-string','text-encoder'],importedStringConstants:'wasm:string-constants'});
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
 assert.equal(imports.filter(i=>i.module==='wasm:text-encoder').length,0,'genuine UTF8 builtins');
 assert.deepEqual(imports.filter(i=>!Object.hasOwn(hosts[i.module]??{},i.name)),[],'no missing/stubbed imports');
 const instance=await WebAssembly.instantiate(module,hosts);
 for(const alias of meta.aliases){const f=meta.exports.find(e=>e.name===`binding-factory:${alias.name}`);assert(f);hosts.env[alias.name].value=instance.exports[f.export](...Array(f.params.length).fill(null));}
 const root=meta.exports[0];instance.exports[root.export](...Array(root.params.length).fill(null));
 const entry=meta.exports.find(e=>e.name==='entry');assert(entry);
 return {run:x=>instance.exports[entry.export](...Array(entry.params.length-1).fill(null),x),stats,bytes:bytes.length};
}

const cases=JSON.parse(readFileSync(join(dir,"LITERAL_RHS_CASES.json")));let checks=0,traps=0;
const signed=n=>{const a=[];for(;;){let b=n&127;n>>=7;const done=(n===0&&!(b&64))||(n===-1&&(b&64));a.push(b|(done?0:128));if(done)return a;}};
const outcome=f=>{try{return {value:f()};}catch(e){return {trap:true};}};
for(const row of cases) {
 const w=await load(row.name);
 const type=[1,0x60,1,0x7f,1,0x7f], fn=[1,0], exp=[1,3,114,117,110,0,0];
 const body=[0,0x20,0,0x41,...signed(row.rhs),row.opcode,...(row.condition?[0x04,0x7f,0x41,1,0x05,0x41,0,0x0b]:[]),0x0b],code=[1,body.length,...body];
 const bytes=Uint8Array.from([0,97,115,109,1,0,0,0,1,type.length,...type,3,fn.length,...fn,7,exp.length,...exp,10,code.length,...code]);
 const {instance}=await WebAssembly.instantiate(bytes,{});
 for(const x of [-2147483648,-2147483647,-65536,-32,-1,0,1,2,31,32,65535,2147483646,2147483647]) {
  const expected=outcome(()=>instance.exports.run(x)),actual=outcome(()=>w.run(x));assert.deepEqual(actual,expected,JSON.stringify({row,x}));checks++;if(expected.trap)traps++;
 }
 for(const x of [null,undefined,true,"1",{},1.5,NaN,Infinity]){assert.throws(()=>w.run(x));checks++;}
}
const result={checks,traps,modules:cases.length,oracle:"Independent minimal raw-i32 Wasm functions executed by V8",node:process.version};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2));console.log(result);
