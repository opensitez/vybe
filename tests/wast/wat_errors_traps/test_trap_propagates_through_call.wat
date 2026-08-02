;; vybe-test: wast/wat_errors_traps/test_trap_propagates_through_call
;; origin: languages/wast/tests/wast/test_wat_errors_traps.rs
;; vybe-test-mode: run-fail

(module
        (import "wasi:logging/logging" "log" (func $log (param i32)))
        (func $boom (result i32) unreachable)
        (func (export "_start") call $boom call $log))
