use std::sync::{Arc,Mutex};
use vybe_ast::Directives;
use vybe_runtime::{Chunk,Op,profile::LanguageProfile,registry::{ParseFn,EmitDispatchFn,emitter_results::{EmitRequest,ResultRepresentation}}};
/// Opt-in contract: standalone complete helper, unchanged existing helper frame,
/// no unit-defined references, no ambient scope/registry-dependent emission;
/// all participating unit emitters preserve append-only chunk-index ownership.
#[derive(Clone,Copy)]pub enum Contract{
 StandaloneExistingFrameAppendOnlyUnit,
 /// Same append-only contextual body ownership, but the only caller emission
 /// is Compiler's checked direct CALL. No arbitrary invoke/caller is run.
 /// Receiver-taking entries require an adapter and are not admitted here;
 /// module_receiver_abi alone does not imply takes_receiver.
 ContextualDirectCallAppendOnlyUnit {
  key:fn(&[Chunk],u32)->Vec<u8>,
  build:fn(&mut Vec<Chunk>,u32)->usize,
 },
 /// Same checked, caller-local-free direct edge, with an exact producer
 /// declaration of the physical Boolean result on every normal return.
 /// The body must also preserve its operation's normal value-position ABI.
 /// No representation is inferred from a name, hint, or signature shape.
 ContextualTypedDirectCallAppendOnlyUnit {
  key:fn(&[Chunk],u32)->Vec<u8>,
  build:fn(&mut Vec<Chunk>,u32)->usize,
  result:ResultRepresentation,
 },
 /// Unit-owned provider whose body depends on an exact, producer-defined
 /// context projection. No cross-unit body cache or relocation is implied.
 /// `build` appends its implementation and returns its owned chunk index;
 /// `invoke` emits the existing caller adaptation against that index.
 /// The standalone factory/caller fields of Definition are unused here.
 ContextualAppendOnlyUnit {
  key:fn(&[Chunk],u32)->Vec<u8>,
  build:fn(&mut Vec<Chunk>,u32)->usize,
  invoke:fn(&mut Vec<Chunk>,usize,usize,u32),
 },
 /// Distinct public adapters may explicitly share one trusted physical body.
 /// Recipe address is authority; a guest-visible name is never a cache key.
 ContextualRecipeAppendOnlyUnit {
  recipe:&'static ContextualRecipe,
  invoke:fn(&mut Vec<Chunk>,usize,usize,u32),
 },
}
pub struct ContextualRecipe {
 pub key:fn(&[Chunk],u32)->Vec<u8>,
 pub build:fn(&mut Vec<Chunk>,u32)->usize,
 pub helper_name:&'static str,
}
/// Compiler callsite frame effect, established by the chosen emission path.
#[derive(Clone,Copy,Debug,PartialEq,Eq)]pub enum FramePolicy { ReserveScope, NoCallerLocals }
/// A successful owned emission's independent frame and physical-result facts.
/// None retains legacy representation provenance, including legacy hints.
#[derive(Clone,Copy,Debug,PartialEq,Eq)]pub struct EmissionReceipt {
 pub frame:FramePolicy,
 pub result:Option<ResultRepresentation>,
}
pub struct Definition{pub contract:Contract,pub operation:&'static str,pub namespace:&'static str,pub version:u32,pub parser:ParseFn,pub producer:EmitDispatchFn,pub arity:u8,pub helper_name:&'static str,pub factory:fn()->Chunk,pub caller:fn(&mut Chunk,usize,u32)}
pub struct Context<'a>{pub profile:&'a Arc<LanguageProfile>,pub directives:&'a Directives,pub receiver:vybe_runtime::chunk::ReceiverAbi,pub own:bool,pub prototype:bool}
struct Entry{definition:usize,profile:Arc<LanguageProfile>,directives:Directives,receiver:vybe_runtime::chunk::ReceiverAbi,own:bool,prototype:bool,body:Arc<Body>}
struct Body{chunk:Chunk,calls:Vec<(u16,u8,u32)>,cacheable:bool}
impl Body{
 fn capture(chunk:Chunk,name:&str)->Self{
  let mut cacheable=chunk.name==name&&chunk.source_path.is_none()&&chunk.direct_defined_calls.is_empty()&&chunk.global_inits.is_empty()&&chunk.constants.iter().all(|v|matches!(v,vybe_runtime::Value::String(_)))&&chunk.types.is_empty()&&chunk.component_funcs.is_empty()&&chunk.component_func_links.is_empty()&&chunk.component_instance_links.is_none()&&chunk.native_segments.is_none()&&chunk.data_segments.is_empty()&&chunk.elem_segments.is_empty()&&chunk.active_elem_segments.is_empty()&&chunk.passive_elem_items.is_empty()&&chunk.memory_min_pages.is_empty()&&chunk.table_min_sizes.is_empty();
  cacheable &= chunk.canon_section.is_empty() && chunk.canon_valtypes.is_empty();
  let mut calls=Vec::new();let mut ip=0;
  while ip<chunk.code.len(){let op=Op::from_code_at(&chunk.code,ip);if op==Op::REF_FUNC{cacheable=false;}let n=op.operand_format().size_in(&chunk.code,ip+4);
   if op==Op::CALL{let i=u16::from_be_bytes(chunk.code[ip+4..ip+6].try_into().expect("registered valid helper"));if i as usize>=chunk.imports.len(){cacheable=false;}calls.push((i,chunk.code[ip+6],chunk.lines[ip]));}ip+=4+n;
  }
  Self{chunk,calls,cacheable}
 }
 fn replay_diagnostics(&self){for&(i,argc,line)in&self.calls{let i=&self.chunk.imports[i as usize];if let Some(declared)=vybe_runtime::host_arity_mismatch(&i.module,&i.name,argc){eprintln!("[vybe] host call arity: {}:{} declares {declared} parameter(s), called with {argc} (line {line})",i.module,i.name);}}}
}
/// Request-owned immutable bodies; no process-global cache or guest indices.
#[derive(Default)]pub struct Pool{entries:Mutex<Vec<Entry>>}
impl Pool{
 fn matches(e:&Entry,d:&Definition,c:&Context<'_>)->bool{e.definition==d as*const Definition as usize&&Arc::ptr_eq(&e.profile,c.profile)&&e.directives==*c.directives&&e.receiver==c.receiver&&e.own==c.own&&e.prototype==c.prototype}
 fn body(&self,d:&Definition,c:&Context<'_>)->Option<(Arc<Body>,bool)>{
  if let Some(body)=self.entries.lock().expect("helper cache poison").iter().find(|e|Self::matches(e,d,c)).map(|e|e.body.clone()){return Some((body,false))}
  // Factory emits diagnostics exactly once for this unit, outside the lock.
  let body=Arc::new(Body::capture((d.factory)(),d.helper_name));
  let mut entries=self.entries.lock().expect("helper cache poison");
  if body.cacheable&&!entries.iter().any(|e|Self::matches(e,d,c)){if entries.len()>=32{entries.remove(0);}entries.push(Entry{definition:d as*const Definition as usize,profile:c.profile.clone(),directives:c.directives.clone(),receiver:c.receiver,own:c.own,prototype:c.prototype,body:body.clone()});}
  Some((body,true))
 }
}
struct ContextualEntry {
 definition:usize,key:Vec<u8>,index:usize,profile:Arc<LanguageProfile>,
 directives:Directives,receiver:vybe_runtime::chunk::ReceiverAbi,own:bool,prototype:bool,
 calls:Vec<(usize,u16,u8,u32)>,
}
#[derive(Default)]pub struct Owner{definitions:&'static[Definition],pool:Option<Arc<Pool>>,indices:Vec<(usize,usize)>,contextual:Vec<ContextualEntry>}
impl Owner{
 pub fn new(definitions:&'static[Definition],pool:Arc<Pool>)->Self{Self{definitions,pool:Some(pool),indices:Vec::new(),contextual:Vec::new()}}
 pub fn emit(&mut self,name:&str,parser:ParseFn,producer:EmitDispatchFn,argc:u8,request:EmitRequest,c:&Context<'_>,chunks:&mut Vec<Chunk>,current:usize,line:u32)->bool{
  self.emit_with_frame_policy(name,parser,producer,argc,request,c,chunks,current,line).is_some()
 }
 pub fn emit_with_frame_policy(&mut self,name:&str,parser:ParseFn,producer:EmitDispatchFn,argc:u8,request:EmitRequest,c:&Context<'_>,chunks:&mut Vec<Chunk>,current:usize,line:u32)->Option<FramePolicy>{
  self.emit_with_result(name,parser,producer,argc,request,c,chunks,current,line).map(|receipt|receipt.frame)
 }
 pub fn emit_with_result(&mut self,name:&str,parser:ParseFn,producer:EmitDispatchFn,argc:u8,request:EmitRequest,c:&Context<'_>,chunks:&mut Vec<Chunk>,current:usize,line:u32)->Option<EmissionReceipt>{
  let Some(pool)=self.pool.as_ref()else{return None};
  let Some((namespace,_))=name.split_once('.')else{return None};
  let Some(i)=self.definitions.iter().position(|d|d.operation==name&&d.namespace==namespace&&d.version>0&&d.arity==argc&&std::ptr::fn_addr_eq(d.parser,parser)&&std::ptr::fn_addr_eq(d.producer,producer))else{return None};let d=&self.definitions[i];
  let result=match d.contract {
   Contract::ContextualTypedDirectCallAppendOnlyUnit {result:result @ ResultRepresentation::BoxedBoolean,..}=>Some(result),
   Contract::ContextualTypedDirectCallAppendOnlyUnit {..}=>return None,
   _=>None,
  };
  if !matches!(request,EmitRequest::Value)
    && !(request==EmitRequest::I32Condition && result.is_some()){return None;}
  let contextual=match d.contract {
   Contract::StandaloneExistingFrameAppendOnlyUnit=>None,
   Contract::ContextualAppendOnlyUnit {key,build,invoke}=>Some((i,key,build,Some(invoke))),
   Contract::ContextualDirectCallAppendOnlyUnit {key,build}
    |Contract::ContextualTypedDirectCallAppendOnlyUnit {key,build,..}=>Some((i,key,build,None)),
   Contract::ContextualRecipeAppendOnlyUnit {recipe,invoke}=>{
    if d.helper_name!=recipe.helper_name{return None}
    // Only explicit registrations for this same producer/version may alias a
    // body. Public arity and invocation remain properties of the selected d.
    let canonical=self.definitions.iter().position(|candidate|
     candidate.namespace==d.namespace&&candidate.version==d.version
     &&candidate.helper_name==recipe.helper_name
     &&std::ptr::fn_addr_eq(candidate.parser,d.parser)
     &&std::ptr::fn_addr_eq(candidate.producer,d.producer)
     &&matches!(candidate.contract,Contract::ContextualRecipeAppendOnlyUnit {recipe:other,..}
      if std::ptr::eq(other,recipe))).expect("selected recipe is registered");
    Some((canonical,recipe.key,recipe.build,Some(invoke)))
   },
  };
  if let Some((definition,key,build,invoke))=contextual {
   let direct=invoke.is_none();
   if direct && current>=chunks.len(){return None;}
   let rollback_chunks=chunks.len();let rollback_entries=self.contextual.len();
   let key=key(chunks,line);
   let existing=self.contextual.iter().position(|entry|
    entry.definition==definition&&entry.key==key&&Arc::ptr_eq(&entry.profile,c.profile)
    &&entry.directives==*c.directives&&entry.receiver==c.receiver
    &&entry.own==c.own&&entry.prototype==c.prototype
    &&chunks.get(entry.index).is_some_and(|chunk|chunk.name==d.helper_name));
   let index=if let Some(existing)=existing {
     let entry=&self.contextual[existing];
     // Reusing construction must not suppress per-emission diagnostics. Keep
     // import identities in the owned chunks and consult live host arities.
     if !direct{replay_contextual_diagnostics(entry,chunks);}
     entry.index
   }else{
    let start=chunks.len();
    // Keep the existing reference encoding's limit and never adopt a helper
    // merely because a guest declaration has a matching name.
    if if direct{u32::try_from(start).is_err()}else{start>u16::MAX as usize}{return None}
    if direct && chunks.iter().enumerate().any(|(index,chunk)|chunk.name==d.helper_name
      && !self.contextual.iter().any(|entry|entry.definition==definition&&entry.index==index)){return None;}
    let index=build(chunks,line);
    if direct {
     if index<start || index>=chunks.len() || chunks[index].name!=d.helper_name {
      chunks.truncate(start);return None;
     }
    }else{
     assert!(index>=start&&index<chunks.len(),"contextual helper must be appended to its unit");
     assert_eq!(chunks[index].name,d.helper_name,"contextual provider identity");
    }
    let mut calls=Vec::new();
    for (chunk_index,chunk) in chunks.iter().enumerate().skip(start) {
     let mut ip=0;
     while ip<chunk.code.len(){
      let op=Op::from_code_at(&chunk.code,ip);
      if op==Op::CALL && !chunk.direct_defined_calls.contains_key(&ip) {
       let import=u16::from_be_bytes(chunk.code[ip+4..ip+6].try_into().expect("valid provider call"));
       assert!((import as usize)<chunk.imports.len(),"contextual helper imports must be owned");
       calls.push((chunk_index,import,chunk.code[ip+6],chunk.lines[ip]));
      }
      ip+=4+op.operand_format().size_in(&chunk.code,ip+4);
     }
    }
    self.contextual.push(ContextualEntry{definition,key,index,profile:c.profile.clone(),
     directives:c.directives.clone(),receiver:c.receiver,own:c.own,prototype:c.prototype,calls});
    index
   };
   if direct {
    // The existing checked emitter writes no caller local or scratch slot.
    // On denial it writes no code/metadata; undo only this append-only build.
    if chunks[index].takes_receiver || !super::functions::try_emit_direct_helper_call(chunks,current,index,argc,line){
     chunks.truncate(rollback_chunks);self.contextual.truncate(rollback_entries);return None;
    }
    if let Some(existing)=existing{replay_contextual_diagnostics(&self.contextual[existing],chunks);}
    return Some(EmissionReceipt{frame:FramePolicy::NoCallerLocals,result});
   }
   invoke.expect("existing contextual adapter")(chunks,current,index,line);
   return Some(EmissionReceipt{frame:FramePolicy::ReserveScope,result:None});
  }
  let helper=if let Some((_,index))=self.indices.iter().find(|(definition,index)|*definition==i&&chunks.get(*index).is_some_and(|c|c.name==d.helper_name)){*index}else{
   // A matching guest-visible name does not prove ownership. A user function
   // may have exactly this spelling; decline rather than calling its body.
   if chunks.iter().any(|c|c.name==d.helper_name){return None}
   let index={
    // Existing ABI represents ref.func targets as u16; do not change its limit.
    if chunks.len()>u16::MAX as usize{return None}
    let Some((body,factory_emitted))=pool.body(d,c)else{return None};
    // The factory has no unit indices. A body outside the registered standalone
    // contract cannot safely be appended even when it is not put in the pool.
    if !body.cacheable{return None}
    if !factory_emitted{body.replay_diagnostics()}
    let index=chunks.len();chunks.push(body.chunk.clone());index
   };self.indices.push((i,index));index
  };
  (d.caller)(&mut chunks[current],helper,line);Some(EmissionReceipt{frame:FramePolicy::ReserveScope,result:None})
 }
}

fn replay_contextual_diagnostics(entry:&ContextualEntry,chunks:&[Chunk]){
 for &(chunk,index,argc,line)in &entry.calls{
  let import=&chunks[chunk].imports[index as usize];
  if let Some(declared)=vybe_runtime::host_arity_mismatch(&import.module,&import.name,argc){
   eprintln!("[vybe] host call arity: {}:{} declares {declared} parameter(s), called with {argc} (line {line})",import.module,import.name);
  }
 }
}

/// Immutable exact-parser selection captured before preparation begins.
/// Bodies are request-owned; target indices are exclusively unit-owned.
#[derive(Clone,Copy)]
pub struct Selection { pub parser:ParseFn, pub definitions:&'static[Definition] }
impl Selection {
 pub fn identity(self)->(usize,usize,usize){(self.parser as usize,self.definitions.as_ptr() as usize,self.definitions.len())}
}
/// Service-level cache choice. Explicit None stays disabled across handoff;
/// an unconfigured request enables a pool only for registered definitions.
#[derive(Clone,Default)]
pub struct Session { pool:Option<Arc<Pool>>, explicit:bool }
impl Session {
 pub(crate) fn has_request_choice(&self)->bool { self.explicit || self.pool.is_some() }
 pub fn set_pool(&mut self,pool:Option<Arc<Pool>>){self.pool=pool;self.explicit=true;}
 pub fn selected(&mut self,selection:Selection)->Option<Selected>{
  if selection.definitions.is_empty(){return None;}
  if self.pool.is_none()&&!self.explicit{self.pool=Some(Arc::new(Pool::default()));}
  self.pool.as_ref().map(|pool|Selected{selection,pool:pool.clone()})
 }
 /// A consumer's explicit choice takes precedence over imported ownership.
 pub fn adopt(&mut self,other:&Self){if !self.explicit{self.pool=other.pool.clone();self.explicit=other.explicit;}}
}
#[derive(Clone)]
pub struct Selected { pub selection:Selection, pool:Arc<Pool> }
impl Selected {pub fn owner(&self)->Owner{Owner::new(self.selection.definitions,self.pool.clone())}}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn canonical_type_owned_body_is_not_context_free() {
        let mut chunk = body();
        crate::primitives::canon::add_typed_stream_import(&mut chunk,
            "stream.drop-readable", vybe_runtime::component::ValType::String,
            Default::default());
        assert!(!Body::capture(chunk, "__registered_template").cacheable);
    }

    #[test]
    fn instance_bound_body_is_not_context_free() {
        assert!(Body::capture(body(), "__registered_template").cacheable);
        let mut chunk = body();
        chunk.component_instance_links = Some(Arc::new(
            vybe_runtime::component_instance_link::InstanceLinkMetadata::default(),
        ));
        assert!(!Body::capture(chunk, "__registered_template").cacheable);
    }

    fn parse(_: &str) -> Result<vybe_ast::Module, String> { unreachable!("identity only") }
    fn other_parse(_: &str) -> Result<vybe_ast::Module, String> { unreachable!("identity only") }
    fn dispatch(_: &str, _: &mut Vec<Chunk>, _: usize, _: u8, _: u32) -> bool { false }
    fn body() -> Chunk {
        let mut chunk = Chunk::new("__registered_template");
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk
    }
    fn foreign_body() -> Chunk {
        let mut chunk = body();
        chunk.emit_op_u16(Op::REF_FUNC, 0, 0);
        chunk.emit(0, 0);
        chunk
    }
    fn caller(chunk: &mut Chunk, index: usize, line: u32) {
        chunk.emit_op_u16(Op::REF_FUNC, index as u16, line);
        chunk.emit(0, line);
        chunk.emit_op_u8_u8(Op::CALL_REF, 0, 1, line);
    }
    static VALID: &[Definition] = &[Definition {
        contract: Contract::StandaloneExistingFrameAppendOnlyUnit,
        operation: "test.value", namespace: "test", version: 1,
        parser: parse, producer: dispatch, arity: 0,
        helper_name: "__registered_template", factory: body, caller,
    }];
    static FOREIGN: &[Definition] = &[Definition {
        contract: Contract::StandaloneExistingFrameAppendOnlyUnit,
        operation: "test.value", namespace: "test", version: 1,
        parser: parse, producer: dispatch, arity: 0,
        helper_name: "__registered_template", factory: foreign_body, caller,
    }];

    fn exercise(definitions: &'static [Definition], collision: bool) {
        let profile = Arc::new(vybe_runtime::profile::parse_profile("[info]\nname='php'\n[compiler]").unwrap());
        let directives = Directives::default();
        let context = Context { profile: &profile, directives: &directives,
            receiver: vybe_runtime::chunk::ReceiverAbi::Ambient, own: false, prototype: false };
        let pool = Arc::new(Pool::default());
        for _ in 0..2 {
            let mut owner = Owner::new(definitions, pool.clone());
            let mut chunks = vec![Chunk::new("caller")];
            if collision {
                let mut guest = Chunk::new("__registered_template");
                guest.emit_i32_const(999, 27);
                chunks.push(guest);
            }
            let before = format!("{chunks:?}");
            assert!(!owner.emit("test.value", other_parse, dispatch, 0, EmitRequest::Value,
                &context, &mut chunks, 0, 13));
            assert_eq!(format!("{chunks:?}"), before, "wrong parser changed caller");
            let admitted = owner.emit("test.value", parse, dispatch, 0, EmitRequest::Value,
                &context, &mut chunks, 0, 13);
            if collision || std::ptr::eq(definitions, FOREIGN) {
                assert!(!admitted);
                assert_eq!(format!("{chunks:?}"), before, "decline must preserve unit and caller");
            } else {
                assert!(admitted);
                assert_eq!(chunks.len(), 2);
                assert!(owner.emit("test.value", parse, dispatch, 0, EmitRequest::Value,
                    &context, &mut chunks, 0, 29));
                assert_eq!(chunks.len(), 2, "same owner must reuse its own helper");
                assert!(chunks[0].lines.contains(&13) && chunks[0].lines.contains(&29));
            }
        }
    }
    #[test]
    fn unowned_guest_name_declines_without_modifying_unit() { exercise(VALID, true); }
    #[test]
    fn unit_defined_factory_reference_declines_without_modifying_unit() { exercise(FOREIGN, false); }
    #[test]
    fn registered_standalone_body_reuses_only_owned_indices() { exercise(VALID, false); }

    fn recursive_key(_: &[Chunk], line: u32) -> Vec<u8> { line.to_le_bytes().to_vec() }
    fn recursive_body(chunks: &mut Vec<Chunk>, line: u32) -> usize {
        let index = chunks.len();
        let mut helper = Chunk::new("__owned_recursive");
        let declaration = helper.clone();
        helper.emit_direct_defined_call(index as u32, &declaration, false, line).unwrap();
        chunks.push(helper);
        index
    }
    fn recursive_invoke(_: &mut Vec<Chunk>, _: usize, _: usize, _: u32) {}
    static RECURSIVE: &[Definition] = &[Definition {
        contract: Contract::ContextualAppendOnlyUnit {
            key: recursive_key, build: recursive_body, invoke: recursive_invoke,
        },
        operation: "test.value", namespace: "test", version: 1,
        parser: parse, producer: dispatch, arity: 0,
        helper_name: "__owned_recursive", factory: body, caller,
    }];
    #[test]
    fn contextual_direct_calls_are_not_host_import_diagnostics() {
        let profile = Arc::new(vybe_runtime::profile::parse_profile("[info]\nname='php'\n[compiler]").unwrap());
        let directives = Directives::default();
        let context = Context { profile: &profile, directives: &directives,
            receiver: Default::default(), own: false, prototype: false };
        let mut owner = Owner::new(RECURSIVE, Arc::new(Pool::default()));
        let mut chunks = vec![Chunk::new("caller")];
        for _ in 0..2 {
            assert!(owner.emit("test.value", parse, dispatch, 0, EmitRequest::Value,
                &context, &mut chunks, 0, 4));
        }
        assert_eq!(chunks.len(), 2);
        assert!(chunks[1].imports.is_empty());
        assert_eq!(chunks[1].direct_defined_calls.len(), 1);
        assert!(owner.contextual[0].calls.is_empty());
    }
    fn no_adapter(_: &mut Chunk,_:usize,_:u32){panic!("direct contract must not invoke arbitrary caller")}
    fn direct_key(_: &[Chunk],_:u32)->Vec<u8>{vec![1]}
    fn direct_body(chunks:&mut Vec<Chunk>,line:u32)->usize{
        let index=chunks.len();let mut c=body();c.direct_entry_has_captures=Some(false);
        let signature=c.clone();c.emit_direct_defined_call(index as u32,&signature,false,line).unwrap();chunks.push(c);index
    }
    fn bad_direct_body(chunks:&mut Vec<Chunk>,line:u32)->usize{let i=direct_body(chunks,line);chunks[i].arity=1;i}
    fn capturing_direct_body(chunks:&mut Vec<Chunk>,line:u32)->usize{let i=direct_body(chunks,line);chunks[i].direct_entry_has_captures=Some(true);i}
    fn receiver_direct_body(chunks:&mut Vec<Chunk>,line:u32)->usize{let i=direct_body(chunks,line);chunks[i].takes_receiver=true;i}
    fn wrong_index_body(chunks:&mut Vec<Chunk>,line:u32)->usize{direct_body(chunks,line);0}
    fn wrong_name_body(chunks:&mut Vec<Chunk>,line:u32)->usize{let i=direct_body(chunks,line);chunks[i].name="wrong_owned_name".into();i}
    static DIRECT:&[Definition]=&[Definition{contract:Contract::ContextualDirectCallAppendOnlyUnit{key:direct_key,build:direct_body},operation:"test.value",namespace:"test",version:1,parser:parse,producer:dispatch,arity:0,helper_name:"__registered_template",factory:body,caller:no_adapter}];
    static BAD_DIRECT:&[Definition]=&[Definition{contract:Contract::ContextualDirectCallAppendOnlyUnit{key:direct_key,build:bad_direct_body},operation:"test.value",namespace:"test",version:1,parser:parse,producer:dispatch,arity:0,helper_name:"__registered_template",factory:body,caller:no_adapter}];
    static CAPTURING_DIRECT:&[Definition]=&[Definition{contract:Contract::ContextualDirectCallAppendOnlyUnit{key:direct_key,build:capturing_direct_body},operation:"test.value",namespace:"test",version:1,parser:parse,producer:dispatch,arity:0,helper_name:"__registered_template",factory:body,caller:no_adapter}];
    static RECEIVER_DIRECT:&[Definition]=&[Definition{contract:Contract::ContextualDirectCallAppendOnlyUnit{key:direct_key,build:receiver_direct_body},operation:"test.value",namespace:"test",version:1,parser:parse,producer:dispatch,arity:0,helper_name:"__registered_template",factory:body,caller:no_adapter}];
    static WRONG_INDEX_DIRECT:&[Definition]=&[Definition{contract:Contract::ContextualDirectCallAppendOnlyUnit{key:direct_key,build:wrong_index_body},operation:"test.value",namespace:"test",version:1,parser:parse,producer:dispatch,arity:0,helper_name:"__registered_template",factory:body,caller:no_adapter}];
    static WRONG_NAME_DIRECT:&[Definition]=&[Definition{contract:Contract::ContextualDirectCallAppendOnlyUnit{key:direct_key,build:wrong_name_body},operation:"test.value",namespace:"test",version:1,parser:parse,producer:dispatch,arity:0,helper_name:"__registered_template",factory:body,caller:no_adapter}];
    fn fixture()->(Arc<LanguageProfile>,Directives){(Arc::new(vybe_runtime::profile::parse_profile("[info]\nname='php'\n[compiler]").unwrap()),Directives::default())}
    #[test]fn direct_context_owned_recursive_body_proves_no_caller_locals(){
        let(p,d)=fixture();let c=Context{profile:&p,directives:&d,receiver:Default::default(),own:false,prototype:false};let mut o=Owner::new(DIRECT,Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];chunks[0].local_count=11;chunks[0].scratch_high_water=17;
        for line in [13,29]{assert_eq!(o.emit_with_frame_policy("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,line),Some(FramePolicy::NoCallerLocals));}
        assert_eq!(chunks.len(),2);assert_eq!(chunks[0].local_count,11);assert_eq!(chunks[0].scratch_high_water,17);assert_eq!(chunks[0].direct_defined_calls.len(),2);assert_eq!(chunks[1].direct_defined_calls.values().next().unwrap().target,1);
        assert!(chunks[0].lines.contains(&13)&&chunks[0].lines.contains(&29));
    }
    #[test]fn direct_denials_preserve_caller_unit_and_ownership(){
        let(p,d)=fixture();let c=Context{profile:&p,directives:&d,receiver:Default::default(),own:false,prototype:false};
        for defs in [DIRECT,BAD_DIRECT,CAPTURING_DIRECT,RECEIVER_DIRECT,WRONG_INDEX_DIRECT,WRONG_NAME_DIRECT]{let mut o=Owner::new(defs,Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];let before=format!("{chunks:?}");
            assert!(o.emit_with_frame_policy("test.value",other_parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,1).is_none());assert!(o.emit_with_frame_policy("test.value",parse,dispatch,1,EmitRequest::Value,&c,&mut chunks,0,1).is_none());assert_eq!(format!("{chunks:?}"),before);
            if !std::ptr::eq(defs,DIRECT){assert!(o.emit_with_frame_policy("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,1).is_none());assert_eq!(format!("{chunks:?}"),before);assert!(o.contextual.is_empty());}
        }
        let mut o=Owner::new(DIRECT,Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller"),Chunk::new("__registered_template")];let before=format!("{chunks:?}");assert!(!o.emit("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,1));assert_eq!(format!("{chunks:?}"),before);assert!(o.contextual.is_empty());
    }
    #[test]fn direct_context_changes_build_new_owned_body_and_wide_indices_use_direct_abi(){
        let(p,d)=fixture();let mut c=Context{profile:&p,directives:&d,receiver:Default::default(),own:false,prototype:false};let mut o=Owner::new(DIRECT,Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];assert!(o.emit("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,1));c.own=true;assert!(o.emit("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,2));assert_eq!(chunks.len(),3);assert_eq!(o.contextual.len(),2);
        let mut o=Owner::new(DIRECT,Arc::new(Pool::default()));let mut chunks=(0..=u16::MAX as usize).map(|_|Chunk::new("padding")).collect::<Vec<_>>();assert!(o.emit("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,3));assert_eq!(chunks[0].direct_defined_calls.values().next().unwrap().target,65536);
    }
    #[test]fn old_contracts_keep_reserve_scope_policy(){
        let(p,d)=fixture();let c=Context{profile:&p,directives:&d,receiver:Default::default(),own:false,prototype:false};
        for defs in [VALID,RECURSIVE]{let mut o=Owner::new(defs,Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];assert_eq!(o.emit_with_frame_policy("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,1),Some(FramePolicy::ReserveScope));}
    }

    #[test]fn direct_no_receiver_contract_accepts_both_module_receiver_abis(){
        let(p,d)=fixture();
        for receiver in [vybe_runtime::chunk::ReceiverAbi::Ambient,vybe_runtime::chunk::ReceiverAbi::Parameter]{
            let c=Context{profile:&p,directives:&d,receiver,own:false,prototype:false};let mut o=Owner::new(DIRECT,Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];chunks[0].module_receiver_abi=receiver;
            assert_eq!(o.emit_with_frame_policy("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,1),Some(FramePolicy::NoCallerLocals));assert!(!chunks[1].takes_receiver);
        }
    }

    fn recipe_invoke_zero(chunks:&mut Vec<Chunk>,current:usize,_:usize,line:u32){chunks[current].emit_i32_const(101,line);}
    fn recipe_invoke_one(chunks:&mut Vec<Chunk>,current:usize,_:usize,line:u32){chunks[current].emit_i32_const(202,line);}
    static SHARED_RECIPE:ContextualRecipe=ContextualRecipe{key:direct_key,build:direct_body,helper_name:"__registered_template"};
    static RECIPE_ADAPTERS:&[Definition]=&[
        Definition{contract:Contract::ContextualRecipeAppendOnlyUnit{recipe:&SHARED_RECIPE,invoke:recipe_invoke_zero},operation:"test.zero",namespace:"test",version:1,parser:parse,producer:dispatch,arity:0,helper_name:"__registered_template",factory:body,caller:no_adapter},
        Definition{contract:Contract::ContextualRecipeAppendOnlyUnit{recipe:&SHARED_RECIPE,invoke:recipe_invoke_one},operation:"test.one",namespace:"test",version:1,parser:parse,producer:dispatch,arity:1,helper_name:"__registered_template",factory:body,caller:no_adapter},
    ];
    #[test]fn canonical_recipe_different_arities_share_only_body_and_keep_adapters(){
        let(p,d)=fixture();let c=Context{profile:&p,directives:&d,receiver:Default::default(),own:false,prototype:false};let mut o=Owner::new(RECIPE_ADAPTERS,Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];
        for (name,argc,line)in [("test.one",1,7),("test.zero",0,9),("test.one",1,11)]{
            assert_eq!(o.emit_with_frame_policy(name,parse,dispatch,argc,EmitRequest::Value,&c,&mut chunks,0,line),Some(FramePolicy::ReserveScope));
        }
        let mut expected=Chunk::new("caller");expected.emit_i32_const(202,7);expected.emit_i32_const(101,9);expected.emit_i32_const(202,11);
        assert_eq!(chunks[0].code,expected.code);assert_eq!(chunks[0].lines,expected.lines);
        assert_eq!(chunks.len(),2);assert_eq!(o.contextual.len(),1);assert_eq!(o.contextual[0].definition,0);
        let before=format!("{chunks:?}");assert!(o.emit_with_frame_policy("test.one",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,13).is_none());assert_eq!(format!("{chunks:?}"),before);
    }

    fn typed_bool_body(chunks:&mut Vec<Chunk>,line:u32)->usize {
        let mut body=Chunk::new("__typed_owned_bool");body.arity=0;body.param_count=0;body.local_count=0;
        body.takes_receiver=false;body.direct_entry_has_captures=Some(false);
        body.emit_bool_const(true,line);body.emit_op(Op::RETURN,line);
        let index=chunks.len();chunks.push(body);index
    }
    fn forbidden_typed_build(_: &mut Vec<Chunk>,_:u32)->usize{panic!("denied request/declaration built a helper")}
    fn typed_definition(result:ResultRepresentation,build:fn(&mut Vec<Chunk>,u32)->usize)->&'static[Definition]{
        Box::leak(vec![Definition{contract:Contract::ContextualTypedDirectCallAppendOnlyUnit{key:direct_key,build,result},operation:"test.value",namespace:"test",version:1,parser:parse,producer:dispatch,arity:0,helper_name:"__typed_owned_bool",factory:body,caller:no_adapter}].into_boxed_slice())
    }
    #[test]fn typed_direct_receipt_is_separate_from_frame_and_request(){
        let(p,d)=fixture();let c=Context{profile:&p,directives:&d,receiver:Default::default(),own:false,prototype:false};
        for(result,build)in[(ResultRepresentation::BoxedBoolean,typed_bool_body as fn(&mut Vec<Chunk>,u32)->usize)]{
            let mut o=Owner::new(typed_definition(result,build),Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];chunks[0].local_count=11;chunks[0].scratch_high_water=17;
            for(request,line)in[(EmitRequest::Value,13),(EmitRequest::I32Condition,29)]{
                assert_eq!(o.emit_with_result("test.value",parse,dispatch,0,request,&c,&mut chunks,0,line),Some(EmissionReceipt{frame:FramePolicy::NoCallerLocals,result:Some(result)}));
            }
            assert_eq!(chunks.len(),2);assert_eq!(chunks[0].local_count,11);assert_eq!(chunks[0].scratch_high_water,17);assert_eq!(chunks[0].direct_defined_calls.len(),2);
            assert!(chunks[0].lines.contains(&13)&&chunks[0].lines.contains(&29));
        }
    }
    #[test]fn typed_direct_unsupported_requests_decline_before_build_without_effects(){
        let(p,d)=fixture();let c=Context{profile:&p,directives:&d,receiver:Default::default(),own:false,prototype:false};
        let mut o=Owner::new(typed_definition(ResultRepresentation::BoxedBoolean,forbidden_typed_build),Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];let before=format!("{chunks:?}");
        for request in [EmitRequest::DiscardedValue,EmitRequest::I32Word,EmitRequest::StringLiteralRhs{constant:0},EmitRequest::NullLiteralRhs,EmitRequest::ScalarLiteralOperand]{
            assert!(o.emit_with_result("test.value",parse,dispatch,0,request,&c,&mut chunks,0,13).is_none());assert_eq!(format!("{chunks:?}"),before);assert!(o.contextual.is_empty());
        }
        for result in [ResultRepresentation::Unknown,ResultRepresentation::RawI32Word,ResultRepresentation::RawI32Boolean]{
            let mut o=Owner::new(typed_definition(result,forbidden_typed_build),Arc::new(Pool::default()));
            for request in [EmitRequest::Value,EmitRequest::I32Condition]{assert!(o.emit_with_result("test.value",parse,dispatch,0,request,&c,&mut chunks,0,13).is_none());assert_eq!(format!("{chunks:?}"),before);assert!(o.contextual.is_empty());}
        }
    }
    #[test]fn old_frame_contracts_do_not_acquire_result_authority_or_condition_admission(){
        let(p,d)=fixture();let c=Context{profile:&p,directives:&d,receiver:Default::default(),own:false,prototype:false};
        for defs in [DIRECT,VALID,RECURSIVE]{let mut o=Owner::new(defs,Arc::new(Pool::default()));let mut chunks=vec![Chunk::new("caller")];let before=format!("{chunks:?}");assert!(o.emit_with_result("test.value",parse,dispatch,0,EmitRequest::I32Condition,&c,&mut chunks,0,1).is_none());assert_eq!(format!("{chunks:?}"),before);let receipt=o.emit_with_result("test.value",parse,dispatch,0,EmitRequest::Value,&c,&mut chunks,0,1).unwrap();assert_eq!(receipt.result,None);assert_eq!(receipt.frame,if std::ptr::eq(defs,DIRECT){FramePolicy::NoCallerLocals}else{FramePolicy::ReserveScope});}
    }

}
