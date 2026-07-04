mod helpers;
mod test_wast_proposals;
mod test_wast_script;
mod test_wat_control_flow;
mod test_wat_execution;
mod test_wat_folded;
mod test_wat_instructions;
mod test_wat_module;
mod test_wat_programs;
mod test_wat_lexical;
mod test_wat_types;
mod test_wat_execution_extended;

// Operations (i32, i64, f32, f64)
mod test_wat_i32_arithmetic;
mod test_wat_i32_bitwise;
mod test_wat_i32_relational;
mod test_wat_i64_arithmetic;
mod test_wat_i64_bitwise;
mod test_wat_i64_relational;
mod test_wat_f32_arithmetic;
mod test_wat_f32_relational;
mod test_wat_f32_rounding;
mod test_wat_f64_arithmetic;
mod test_wat_f64_relational;
mod test_wat_f64_rounding;

// Wasm GC (Structs & Arrays)
mod test_wat_struct_new;
mod test_wat_struct_get;
mod test_wat_struct_set;
mod test_wat_array_new;
mod test_wat_array_get_set;
mod test_wat_ref_cast;
mod test_wat_ref_null;

// Functions & Control Flow
mod test_wat_func_params_returns;
mod test_wat_func_locals;
mod test_wat_call_direct;
mod test_wat_call_indirect;
mod test_wat_return_call;
mod test_wat_block;
mod test_wat_loop;
mod test_wat_if_else;
mod test_wat_br_table;

// Memory & Globals
mod test_wat_memory_load;
mod test_wat_memory_store;
mod test_wat_memory_ops;
mod test_wat_globals_mut;
mod test_wat_globals_const;
