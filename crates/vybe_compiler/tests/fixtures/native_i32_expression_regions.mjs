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
 return {run:(...xs)=>instance.exports[entry.export](...Array(entry.params.length-xs.length).fill(null),...xs),stats,bytes:bytes.length};
}

const cases=JSON.parse(readFileSync(join(dir,"I32_REGION_CASES.json")));let checks=0,traps=0;
const leb=n=>{const a=[];do{let b=n&127;n>>>=7;a.push(b|(n?128:0));}while(n);return a;};
const signed=n=>{const a=[];for(;;){let b=n&127;n>>=7;const done=(n===0&&!(b&64))||(n===-1&&(b&64));a.push(b|(done?0:128));if(done)return a;}};
const section=(id,data)=>[id,...leb(data.length),...data];
const outcome=f=>{try{return {value:f()};}catch(e){return {name:e.constructor.name,message:e.message};}};
const pool=[-2147483648,-65536,-32,-1,0,1,2,31,65535,2147483647];
for(const row of cases) {
 const entries=await Promise.all(row.names.map(load));
 const instructions=[];
 for(const [kind,value] of row.tokens){if(kind==="get")instructions.push(0x20,...leb(value));else if(kind==="set")instructions.push(0x21,...leb(value));else if(kind==="const")instructions.push(0x41,...signed(value));else if(kind==="op")instructions.push(value);else if(kind!=="call")throw Error(kind);}
 if(row.condition)instructions.push(0x04,0x7f,0x41,1,0x05,0x41,0,0x0b);
 const body=[1,1,0x7f,...instructions,0x0b];
 const bytes=Uint8Array.from([0,97,115,109,1,0,0,0,...section(1,[1,0x60,3,0x7f,0x7f,0x7f,1,0x7f]),...section(3,[1,0]),...section(7,[1,3,114,117,110,0,0]),...section(10,[1,...leb(body.length),...body])]);
 const {instance}=await WebAssembly.instantiate(bytes,{});
 for(let n=0;n<100;n++){
  const args=[pool[n%10],pool[(n/10|0)%10],pool[(n*7+n/10|0)%10]];
  const expected=outcome(()=>instance.exports.run(...args));
  for(const w of entries){const actual=outcome(()=>w.run(...args));assert.deepEqual(actual,expected,JSON.stringify({row:row.names,args}));checks++;if(expected.name)traps++;}
 }
 for(let slot=0;slot<3;slot++)for(const bad of [null,undefined,true,"1",{},1.5,NaN,Infinity,-0]){
  const args=[1,1,1];args[slot]=bad;
  assert.deepEqual(outcome(()=>entries[1].run(...args)),outcome(()=>entries[0].run(...args)),row.names.join("/"));checks++;
 }
}
const result={checks,traps,modules:cases.length*2,oracleModules:cases.length,node:process.version,scope:"Numeric regions versus independent raw Wasm; additional exact error-order comparisons against boxing barriers, stores, nested operands and conditional consumers"};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2));console.log(result);
