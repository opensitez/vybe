pub mod bigint;
pub mod capabilities;
pub mod chunk;
pub mod js_builtins; // wasm:js-* CG proposals (dispatcher)
pub mod js_primitive_builtins; // wasm:js-{number,boolean,undefined,symbol,bigint}
pub mod js_string_builtins; // wasm:js-string (merged js-string-builtins)
pub mod opcode;
pub mod type_recorder;
pub mod value;
pub mod vm;
// `impl VM` partials — extracted from vm.rs for readability. Each file is
// its own `impl VM { ... }` block operating on the same struct defined in
// vm.rs. Private to the crate; external consumers keep using `VM::*`.
pub(crate) mod calls;
pub mod cm_task;
pub mod debug;
pub mod debugger;
pub(crate) mod dispatch;
pub mod error;
pub mod event_loop;
pub mod heap;
pub mod fiber;
pub mod handle_table;
pub(crate) mod jspi;
pub mod module_record;
pub mod shared_memory;
pub(crate) mod simd;
pub(crate) mod threads;
pub mod typedef;
pub(crate) mod upvalues;
pub mod waitable;

pub use bigint::{BigIntRef, BigIntVal};
pub use chunk::{Chunk, Import};
pub use debugger::{DebugCommand, DebugEvent, DebugRequest, DebugResponse, Debugger};
pub use error::VMError;
pub use event_loop::EventLoop;
pub use module_record::{ExportEntry, ModuleKind, ModuleRecord, ModuleRequest, ModuleStatus};
pub use opcode::Op;
pub use typedef::{FieldDef, Method, ResourceTable, TypeDef, TypeRegistry};
pub use value::Value;
pub use vm::{HostContext, HostFn, ImportTarget, VM, VmSnapshot};
pub mod component;
pub mod component_model;
pub mod project;
pub use component::{
    BinaryLoader, Component, ExportImpl, FuncSig, ImportPolicy, Interface, Language, LinkResult,
    Linker, ModuleExport, ModuleResolver, ResolvedModule, ValType, register_binary_loader,
};
pub use project::ProjectConfig;
