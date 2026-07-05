//! Per-proposal modules — each file implements one WebAssembly proposal
//! end-to-end: the imports it declares, the opcodes it emits, the custom
//! sections it produces.

pub mod bulk_memory;
pub mod compilation_hints;
pub mod esm_integration;
pub mod exception_handling;
pub mod extended_name_section;
pub mod gc;
pub mod jspi;
pub mod multi_value;
pub mod nontrapping_float_to_int;
pub mod reference_types;
pub mod simd;
pub mod stack_switching;
pub mod tail_call;
pub mod threads;
