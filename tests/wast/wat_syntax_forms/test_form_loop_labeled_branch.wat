;; vybe-test: wast/wat_syntax_forms/test_form_loop_labeled_branch
;; origin: languages/wast/tests/wast/test_wat_syntax_forms.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (func (export "_start") (local $i i32) (local $s i32)
  (block $b (loop $l
    local.get $i i32.const 5 i32.eq br_if $b
    local.get $s local.get $i i32.add local.set $s
    local.get $i i32.const 1 i32.add local.set $i
    br $l))
  local.get $s call $log)
)
