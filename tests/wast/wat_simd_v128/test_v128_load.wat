;; vybe-test: wast/wat_simd_v128/test_v128_load
;; origin: languages/wast/tests/wast/test_wat_simd_v128.rs

(module
  (import "wasi:logging/logging" "log" (func $log (param i32)))
  (import "wasi:logging/logging" "log" (func $log_i64 (param i64)))
  (import "wasi:logging/logging" "log" (func $log_f32 (param f32)))
  (import "wasi:logging/logging" "log" (func $log_f64 (param f64)))
  (memory 1)
(data (i32.const 0) "\04\03\02\01\08\07\06\05\0c\0b\0a\09\10\0f\0e\0d")
(func (export "_start")
  i32.const 0
  v128.load
  i32x4.extract_lane 3
  call $log
)
)
