use crate::wat_exec;

wat_exec! {
    test_simd_i32x4_add => { r#"
(func (export "_start")
  v128.const i32x4 10 20 30 40
  v128.const i32x4 5 15 25 35
  i32x4.add
  i32x4.extract_lane 2
  call $log
)
"#, "55" },

    test_simd_i32x4_sub => { r#"
(func (export "_start")
  v128.const i32x4 10 20 30 40
  v128.const i32x4 5 25 25 35
  i32x4.sub
  i32x4.extract_lane 1
  call $log
)
"#, "-5" },

    test_simd_i32x4_mul => { r#"
(func (export "_start")
  v128.const i32x4 10 20 30 40
  v128.const i32x4 5 25 25 35
  i32x4.mul
  i32x4.extract_lane 0
  call $log
)
"#, "50" },

    test_simd_f32x4_add => { r#"
(func (export "_start")
  v128.const f32x4 1.5 2.5 3.5 4.5
  v128.const f32x4 0.5 1.5 2.5 3.5
  f32x4.add
  f32x4.extract_lane 3
  call $log_f32
)
"#, "8.0" },

    test_simd_f32x4_div => { r#"
(func (export "_start")
  v128.const f32x4 1.5 2.5 10.0 4.5
  v128.const f32x4 0.5 1.5 2.0 3.5
  f32x4.div
  f32x4.extract_lane 2
  call $log_f32
)
"#, "5.0" },

    test_simd_i16x8_add_sat_s => { r#"
(func (export "_start")
  v128.const i16x8 32767 0 0 0 0 0 0 0
  v128.const i16x8 10 0 0 0 0 0 0 0
  i16x8.add_sat_s
  i16x8.extract_lane_s 0
  call $log
)
"#, "32767" },

    test_simd_i16x8_add_sat_u => { r#"
(func (export "_start")
  v128.const i16x8 -1 0 0 0 0 0 0 0 ;; 65535
  v128.const i16x8 10 0 0 0 0 0 0 0
  i16x8.add_sat_u
  i16x8.extract_lane_u 0
  call $log
)
"#, "65535" },

    test_simd_i8x16_sub_sat_s => { r#"
(func (export "_start")
  v128.const i8x16 -128 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
  v128.const i8x16 10 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
  i8x16.sub_sat_s
  i8x16.extract_lane_s 0
  call $log
)
"#, "-128" },

    test_simd_i8x16_sub_sat_u => { r#"
(func (export "_start")
  v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
  v128.const i8x16 10 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
  i8x16.sub_sat_u
  i8x16.extract_lane_u 0
  call $log
)
"#, "0" },

    test_simd_i32x4_eq => { r#"
(func (export "_start")
  v128.const i32x4 10 20 30 40
  v128.const i32x4 5 20 25 35
  i32x4.eq
  i32x4.extract_lane 1
  call $log
)
"#, "-1" }, // all ones

    test_simd_i32x4_ne => { r#"
(func (export "_start")
  v128.const i32x4 10 20 30 40
  v128.const i32x4 5 20 25 35
  i32x4.ne
  i32x4.extract_lane 1
  call $log
)
"#, "0" } // zero
}
