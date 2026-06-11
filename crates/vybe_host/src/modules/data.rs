//! System.Data — retired.
//!
//! All vybe:data host fns replaced by compile-time adapters in
//! `emitter/dotnet/core/datatable_adapter.rs`:
//!   DataTable/DataSet constructors → inline STRUCT_NEW bytecode
//!   DataTable.NewRow/AddRow/Select  → inline bytecode via common dispatch
//!   DataSet.Tables                  → inline STRUCT_GET "tables"
//!   DataRow.Item/IsNull             → ecma:object.get + null check
//!   DBNull.Value                    → Op::NULL

use vybe_bytecode::VM;

pub fn register(_vm: &mut VM) {}
