//! Multi-memory proposal — a module may declare several memories; loads and
//! stores carry a memory index.
use crate::wat_exec;

wat_exec! {
    test_two_memories_independent => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory $a 1) (memory $b 1)
        (func (export "_start")
          i32.const 0 i32.const 111 i32.store 0
          i32.const 0 i32.const 222 i32.store 1
          i32.const 0 i32.load 1 call $log))"#, "222" },
    test_first_memory_unaffected_by_second => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory $a 1) (memory $b 1)
        (func (export "_start")
          i32.const 0 i32.const 111 i32.store 0
          i32.const 0 i32.const 222 i32.store 1
          i32.const 0 i32.load 0 call $log))"#, "111" },
    test_memory_size_of_second => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory $a 1) (memory $b 3)
        (func (export "_start") memory.size 1 call $log))"#, "3" },
    test_named_memory_reference => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory $data 1) (memory $scratch 1)
        (func (export "_start")
          i32.const 4 i32.const 999 i32.store $scratch
          i32.const 4 i32.load $scratch call $log))"#, "999" },
    test_copy_between_memories => { r#"(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (memory $a 1) (memory $b 1)
        (func (export "_start")
          i32.const 0 i32.const 42 i32.store 0
          i32.const 0 i32.const 0 i32.const 4 memory.copy 1 0
          i32.const 0 i32.load 1 call $log))"#, "42" },
}
