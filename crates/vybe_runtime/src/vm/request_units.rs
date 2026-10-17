//! Managed source-unit residency. Native instance storage keeps the raw path.
use super::*;
use std::sync::{Arc,Weak};

#[derive(Clone,Debug)]
pub struct LoadedRequestUnit(Arc<LoadedUnit>);
#[derive(Debug)]
struct LoadedUnit {
    epoch:Weak<()>, first:usize, end:usize,
    native_imports:Vec<Option<NativeCallSiteTarget>>,
}
#[derive(Clone,Copy,Debug,PartialEq,Eq)]
pub enum RequestUnitDenial { Empty, NativeInstance, Component, Canonical, Storage, MutableConstant }

fn immutable_value(value:&Value)->bool {
    matches!(value,Value::Null|Value::Undefined|Value::Bool(_)|Value::I32(_)|Value::I64(_)
        |Value::F32(_)|Value::F64(_)|Value::String(_)|Value::V128(_))
}
fn immutable_initializer(expr:&crate::chunk::ConstExpr)->bool {
    use crate::chunk::ConstExpr;
    match expr {
        ConstExpr::Value(value)=>immutable_value(value),
        ConstExpr::Add(a,b)|ConstExpr::Mul(a,b)=>immutable_initializer(a)&&immutable_initializer(b),
        ConstExpr::RefI31(value)=>immutable_initializer(value),
        ConstExpr::StructNew{fields,..}=>fields.iter().all(immutable_initializer),
        ConstExpr::ArrayNewDefault{len,..}=>immutable_initializer(len),
        ConstExpr::ArrayNew{value,len,..}=>immutable_initializer(value)&&immutable_initializer(len),
        _=>true,
    }
}
fn native_projection(target:&ImportTarget)->Option<NativeCallSiteTarget> {
    NativeCallSiteTarget::from_resolved(&ResolvedCallTarget::from_import_target_ref(target))
}
impl LoadedRequestUnit {
    pub fn chunk_range(&self)->std::ops::Range<usize>{self.0.first..self.0.end}
    pub fn admission(chunks:&[Chunk])->Result<(),RequestUnitDenial>{
        if chunks.is_empty(){return Err(RequestUnitDenial::Empty)}
        for c in chunks {
            if c.native_global_import_count.is_some()||c.native_segments.is_some(){return Err(RequestUnitDenial::NativeInstance)}
            if !c.component_funcs.is_empty()||!c.component_func_links.is_empty()||c.component_instance_links.is_some(){return Err(RequestUnitDenial::Component)}
            if !c.canon_section.is_empty()||!c.canon_functypes.is_empty()||!c.canon_valtypes.is_empty()
                ||c.imports.iter().any(|i|i.module=="canon"){return Err(RequestUnitDenial::Canonical)}
            if !c.memory_min_pages.is_empty()||!c.memory_max_pages.is_empty()||!c.memory_is_64.is_empty()
                ||!c.table_min_sizes.is_empty()||!c.table_max_sizes.is_empty()||!c.table_is_64.is_empty()
                ||!c.table_element_types.is_empty()||!c.table_is_shared.is_empty()
                ||!c.data_segments.is_empty()||!c.elem_segments.is_empty()||!c.passive_elem_items.is_empty()
                ||!c.active_data_segments.is_empty()||!c.active_elem_segments.is_empty(){return Err(RequestUnitDenial::Storage)}
            if !c.constants.iter().all(immutable_value)||!c.global_inits.iter().all(|g|immutable_initializer(&g.init)){
                return Err(RequestUnitDenial::MutableConstant)
            }
        }
        Ok(())
    }
    pub fn accepts(chunks:&[Chunk])->bool{Self::admission(chunks).is_ok()}
}
impl VM {
    pub fn run_linked_nested_request(&mut self,chunks:Vec<Chunk>,imports:Vec<ImportTarget>)
        ->(Result<Value,VMError>,Option<LoadedRequestUnit>){
        let mut unit=None;let result=self.run_linked_impl_recorded(chunks,imports,true,None,Some(&mut unit));(result,unit)
    }
    pub fn run_linked_nested_prepared_request(&mut self,code:crate::PreparedLinkedCode,imports:Vec<ImportTarget>)
        ->(Result<Value,VMError>,Option<LoadedRequestUnit>){
        let(chunks,metadata)=code.into_parts();let mut unit=None;
        let result=self.run_linked_impl_recorded(chunks,imports,true,Some(metadata),Some(&mut unit));(result,unit)
    }
    pub(super) fn seal_request_unit(&self,first:usize,end:usize)->Option<LoadedRequestUnit>{
        let script=self.chunks.get(first)?;self.chunks.get(first..end)?;
        if first>=end||script.imports.len()!=self.import_table.len(){return None}
        Some(LoadedRequestUnit(Arc::new(LoadedUnit{epoch:Arc::downgrade(&self.request_code_epoch),first,end,
            native_imports:self.import_table.iter().map(native_projection).collect()})))
    }
    pub fn request_unit_valid(&self,unit:&LoadedRequestUnit)->bool{
        unit.0.epoch.upgrade().is_some_and(|e|Arc::ptr_eq(&e,&self.request_code_epoch))
            &&unit.0.first<unit.0.end&&self.chunks.get(unit.chunk_range()).is_some()
    }
    pub fn request_unit_imports(&self,unit:&LoadedRequestUnit)->Result<&[crate::chunk::Import],VMError>{
        if !self.request_unit_valid(unit){return Err(VMError::new("stale request code unit"))}
        Ok(&self.chunks[unit.0.first].imports)
    }
    /// Pure preflight. A denial can take the raw loader before guest execution.
    /// Fresh resolution is supplied exactly as for ordinary run_linked_nested.
    pub fn request_unit_accepts_imports(&self,unit:&LoadedRequestUnit,imports:&[ImportTarget])->bool{
        self.request_unit_valid(unit)&&imports.len()==unit.0.native_imports.len()
            &&imports.iter().map(native_projection).eq(unit.0.native_imports.iter().copied())
    }
    pub fn invalidate_request_units(&mut self){self.request_code_epoch=Arc::new(())}
    pub fn run_resident_nested(&mut self,unit:&LoadedRequestUnit,imports:Vec<ImportTarget>)->Result<Value,VMError>{
        if !self.request_unit_accepts_imports(unit,&imports){return Err(VMError::new("stale request code unit or changed native import binding"))}
        let saved_imports=Some(std::mem::take(&mut self.import_table));
        let saved_resolved=Some(std::mem::take(&mut self.resolved_import_table));
        self.chunk_import_cache.clear();self.chunk_call_target_cache.clear();
        self.install_linked_import_context(unit.0.first,imports);
        let result=(||{
            self.bind_imported_globals_in_range(unit.chunk_range());
            self.initialize_global_inits(unit.0.first)?;
            self.execute_linked_script_entry(unit.0.first,true)
        })();
        self.restore_linked_import_context(saved_imports,saved_resolved);result
    }
    pub(super) fn execute_linked_script_entry(&mut self,script_idx:usize,nested:bool)->Result<Value,VMError>{
        let saved_frame_depth=self.frames.len();let saved_stack_len=self.stack.len();
        self.ensure_execution_fast_paths(script_idx);
        self.frames.push(CallFrame{chunk_index:script_idx,ip:0,base:self.stack.len(),
            local_frame_size:self.chunks[script_idx].local_count as usize,label_base:self.label_stack.len(),upvalues:Vec::new()});
        self.stack.resize(self.stack.len()+self.chunks[script_idx].local_count as usize,Value::Null);
        if nested{let result=self.execute_until(saved_frame_depth+1);self.stack.truncate(saved_stack_len);result}
        else{let result=self.execute();self.pending_exit=false;result}
    }
    pub(super) fn restore_linked_import_context(&mut self,imports:Option<Vec<ImportTarget>>,resolved:Option<Vec<ResolvedCallTarget>>){
        if let Some(imports)=imports{self.import_table=imports;self.chunk_import_cache.clear();self.chunk_call_target_cache.clear()}
        if let Some(resolved)=resolved{self.resolved_import_table=resolved}
    }
}

#[cfg(any(test,v6_resident_gate))]
#[path="request_unit_tests.rs"]
mod tests;
#[cfg(v6_resident_gate)]
pub fn run_resident_native_gate(){tests::run_all()}

impl VM {
    pub(super) fn install_linked_import_context(&mut self,script_idx:usize,resolved_imports:Vec<ImportTarget>){
        for target in resolved_imports {
            match target {
                ImportTarget::ChunkFn { chunk_index, arity } => {
                    self.import_table.push(ImportTarget::ChunkFn {
                        chunk_index: chunk_index + script_idx,
                        arity,
                    });
                }
                ImportTarget::Host(idx) => {
                    self.import_table.push(ImportTarget::Host(idx));
                }
                ImportTarget::StdlibRedirect(name) => {
                    // Try to resolve against host registry first
                    let key_candidates = [("wasi:cli".to_string(), name.clone())];
                    let mut resolved = false;
                    for key in &key_candidates {
                        if let Some(idx) = self.resolve_host_function_index(&key.0, &key.1) {
                            self.import_table.push(ImportTarget::Host(idx));
                            resolved = true;
                            break;
                        }
                    }
                    if !resolved {
                        self.import_table.push(ImportTarget::StdlibRedirect(name));
                    }
                }
                ImportTarget::JspiSuspend => {
                    self.import_table.push(ImportTarget::JspiSuspend);
                }
                ImportTarget::JspiSuspendEager => {
                    self.import_table.push(ImportTarget::JspiSuspendEager);
                }
                ImportTarget::WasiThreadSpawn => {
                    self.import_table.push(ImportTarget::WasiThreadSpawn);
                }
                ImportTarget::JsNumber(builtin) => {
                    self.import_table.push(ImportTarget::JsNumber(builtin));
                }
                ImportTarget::JsBoolean(builtin) => {
                    self.import_table.push(ImportTarget::JsBoolean(builtin));
                }
                ImportTarget::JsUndefined(builtin) => {
                    self.import_table.push(ImportTarget::JsUndefined(builtin));
                }
                ImportTarget::JsString(builtin) => {
                    self.import_table.push(ImportTarget::JsString(builtin));
                }
                ImportTarget::EcmaNumber(builtin, host_idx) => {
                    self.import_table
                        .push(ImportTarget::EcmaNumber(builtin, host_idx));
                }
                ImportTarget::EcmaBoolean(builtin, host_idx) => {
                    self.import_table
                        .push(ImportTarget::EcmaBoolean(builtin, host_idx));
                }
                ImportTarget::EcmaObject(builtin, host_idx) => {
                    self.import_table
                        .push(ImportTarget::EcmaObject(builtin, host_idx));
                }
                ImportTarget::EcmaArray(builtin, host_idx) => {
                    self.import_table
                        .push(ImportTarget::EcmaArray(builtin, host_idx));
                }
                ImportTarget::JspiYield => {
                    self.import_table.push(ImportTarget::JspiYield);
                }
                ImportTarget::StringConst(s) => {
                    self.import_table.push(ImportTarget::StringConst(s));
                }
                ImportTarget::Canon { builtin, type_idx, component_result } => {
                    self.import_table.push(ImportTarget::Canon {
                        builtin,
                        type_idx,
                        component_result,
                    });
                }
            }
        }
        self.rebuild_resolved_import_table();
    }
}
