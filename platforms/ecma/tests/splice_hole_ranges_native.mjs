import assert from 'node:assert/strict';
import {readFileSync,writeFileSync} from 'node:fs';
const queries=JSON.parse(readFileSync(process.argv[2])).queries;
function snapshot(a){return {values:Array.from({length:a.length},(_,i)=>Object.hasOwn(a,i)?a[i]:null),holes:Array.from({length:a.length},(_,i)=>i).filter(i=>!Object.hasOwn(a,i))};}
let checks=0;
for(const q of queries){const a=Array.from({length:q.len},(_,i)=>i);for(const h of q.holes)delete a[h];const deleted=a.splice(q.start,q.delete,...Array.from({length:q.insert},(_,i)=>1000+i));assert.deepEqual(q.after,snapshot(a),JSON.stringify(q));assert.deepEqual(q.removed,snapshot(deleted),JSON.stringify(q));checks+=2;}
const result={checks,cases:queries.length,node:process.version,scope:'Actual Rust splice provider output versus unmodified Node Array.splice; values and own-index hole presence; this is a host-adapter oracle, not exported Wasm execution or timing'};
if(process.argv[3])writeFileSync(process.argv[3],JSON.stringify(result,null,2)+'\n');console.log(result);
