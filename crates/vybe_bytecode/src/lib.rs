pub mod opcode;
pub mod value;
pub mod chunk;
pub mod vm;
pub mod type_recorder;
// `impl VM` partials — extracted from vm.rs for readability. Each file is
// its own `impl VM { ... }` block operating on the same struct defined in
// vm.rs. Private to the crate; external consumers keep using `VM::*`.
pub(crate) mod dispatch;
pub(crate) mod calls;
pub(crate) mod upvalues;
pub(crate) mod jspi;
pub(crate) mod simd;
pub mod error;
pub mod debug;
pub mod fiber;
pub mod event_loop;
pub mod typedef;
pub mod shared_memory;
pub mod module_record;

pub use opcode::Op;
pub use value::Value;
pub use chunk::{Chunk, Import};
pub use vm::{VM, HostFn, HostContext, ImportTarget};
pub use error::VMError;
pub use event_loop::EventLoop;
pub use typedef::{TypeDef, TypeRegistry, Method, FieldDef, ResourceTable};
pub use module_record::{ModuleRecord, ModuleKind, ModuleStatus, ExportEntry, ModuleRequest};
pub mod component;
pub mod component_model;
pub mod project;
pub use component::{Component, Linker, LinkResult, Interface, FuncSig, ValType, Language, ExportImpl, ModuleResolver, ResolvedModule, ModuleExport, ImportPolicy};
pub use project::ProjectConfig;
pub mod wasm;
