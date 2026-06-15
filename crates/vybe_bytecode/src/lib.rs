pub mod chunk;
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
pub(crate) mod dispatch;
pub mod error;
pub mod event_loop;
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

pub use chunk::{Chunk, Import};
pub use error::VMError;

/// Suspend tag reserved for JS `await`.
///
/// `await` lowers to the spec stack-switching `suspend` instruction (JSPI is
/// the stack-switching proposal applied to JS Promises). To keep `await`
/// distinct from a generator `yield` — both of which use `suspend` — `await`
/// carries this dedicated tag. The VM's `SUSPEND` handler routes this tag to
/// the Promise-await behaviour (settle/throw/suspend-on-pending) regardless of
/// any active generator continuation, while tag 0 stays generator `yield`.
pub const AWAIT_SUSPEND_TAG: u16 = 0xFFFF;
pub use event_loop::EventLoop;
pub use module_record::{ExportEntry, ModuleKind, ModuleRecord, ModuleRequest, ModuleStatus};
pub use opcode::Op;
pub use typedef::{FieldDef, Method, ResourceTable, TypeDef, TypeRegistry};
pub use value::Value;
pub use vm::{HostContext, HostFn, ImportTarget, VM};
pub mod component;
pub mod component_model;
pub mod project;
pub use component::{
    Component, ExportImpl, FuncSig, ImportPolicy, Interface, Language, LinkResult, Linker,
    ModuleExport, ModuleResolver, ResolvedModule, ValType,
};
pub use project::ProjectConfig;
pub mod wasm;
