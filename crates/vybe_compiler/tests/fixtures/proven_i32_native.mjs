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
 env:{}
 };
 Object.assign(hosts['ecma:array'],{isArray:Array.isArray,get:(a,k)=>Reflect.get(Object(a),String(k),a),set:(a,k,v)=>{Reflect.set(a,String(k),v,a);},splice:(a,start,count,...items)=>Array.prototype.splice.call(a,start,count,...items)});
 Object.assign(hosts['ecma:object'],{keys:Object.keys,assign:Object.assign});
 hosts['ecma:number']={isInteger:Number.isInteger};
 hosts['ecma:map']={size:x=>x instanceof Map?x.size:0};
 hosts['ecma:bigint']={add:(a,b)=>a+b,lt:(a,b)=>a<b,gt:(a,b)=>a>b};
 hosts['ecma:string']={String:String,padEnd:(s,n,p)=>String.prototype.padEnd.call(s,n,p)};
 hosts['ecma:reflect']={isCallable:x=>typeof x==='function'};
 hosts['vybe:js-array'].fill=(a,v,start,end)=>Array.prototype.fill.call(a,v,start,end);
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
let checks=0;const rows=[];
for(let kind=0;kind<6;kind++){
 const ws=await Promise.all([false,true].map(fast=>load(`proven-i32-${kind}-${fast}`)));
 const values=[0,7,.5,NaN,Infinity,-0,null,undefined,true,false,'','7',0n,[],{},Symbol('x')];
 if(kind<3)for(const x of values){const out=ws.map(w=>w.run(x));assert(Object.is(out[0],out[1]));if(kind===0)assert.equal(out[1],typeof x==='number'?+!Number.isInteger(x):false);if(kind===2)assert.equal(out[1],typeof x==='number');checks+=2;}
 if(kind===3)for(const input of [[7,2],[true],['s'],[]]){const outputs=ws.map(w=>w.run([...input]));assert.deepEqual(outputs[0],outputs[1]);const fill=typeof input[0]==='boolean'?false:typeof input[0]==='number'?0:null;assert.deepEqual(outputs[1],[...input,...Array(4-input.length).fill(fill)]);checks+=2;}
 if(kind===4||kind===5){for(const [x,y]of [['abcd','Q'],['abcd','LONG'],['é😀z','x'],['😀z',7],['abcd',null]]){const out=ws.map(w=>w.run(x,y));assert.deepEqual(out[0],out[1]);const fragment=String(y).substring(0,2).padEnd(2,' ');assert.equal(out[1],x.substring(0,1)+fragment+x.substring(3));checks+=2;}
 if(kind===4)for(const replacement of ['Q',[8]]){const out=ws.map(w=>w.run([0,1,2,3],Array.isArray(replacement)?[...replacement]:replacement));assert.deepEqual(out[0],out[1]);assert.deepEqual(out[1],Array.isArray(replacement)?[0,8,3]:[0,'Q','Q',3]);checks+=2;}}
 rows.push({kind,bytes:ws.map(w=>w.bytes)});
}
for(const kind of [0,1]){const ws=await Promise.all([false,true].map(fast=>load(`proven-i32-${kind}-${fast}`)));for(const w of ws){let calls=0;const x=new Proxy({},{get(){calls++;throw Error('unexpected getter');}});w.run(x);assert.equal(calls,0);checks++;}}
// Conversion of a replacement object remains at the original call point.
for(const kind of [4,5]){const ws=await Promise.all([false,true].map(fast=>load(`proven-i32-${kind}-${fast}`)));const outcomes=ws.map(w=>{const trace=[];const y={[Symbol.toPrimitive](hint){trace.push(hint);return 'Z';}};return {value:w.run('abcd',y),trace};});assert.deepEqual(outcomes[0],outcomes[1]);assert.deepEqual(outcomes[1],{value:'aZ d',trace:['string']});checks+=2;}
const result={checks,modules:12,node:process.version,rows,scope:'Compiler-produced old/new guest bodies with appended diagnostic RefFunc factories for imported bundled helpers; full native validation occurs before and after diagnostic exports. Real js-string builtins. Primitive type inputs, legacy result tags, dense resize/fill, typed/dynamic UTF16 string slice assignment, dynamic array splice, replacement coercion order and pure type-test proxy behavior. Constructor spread missing/null controls covered by VM/native validation only; no generator execution or native speed claim.'};if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
