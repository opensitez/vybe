//! SIMD v128 bitwise logic — operate on all 128 bits regardless of lane shape.
use crate::wat_exec;

wat_exec! {
    test_v128_and => { r#"(func (export "_start")
        v128.const i32x4 0xFF 0 0 0 v128.const i32x4 0x0F 0 0 0
        v128.and i32x4.extract_lane 0 call $log)"#, "15" },
    test_v128_or => { r#"(func (export "_start")
        v128.const i32x4 0xF0 0 0 0 v128.const i32x4 0x0F 0 0 0
        v128.or i32x4.extract_lane 0 call $log)"#, "255" },
    test_v128_xor => { r#"(func (export "_start")
        v128.const i32x4 0xFF 0 0 0 v128.const i32x4 0x0F 0 0 0
        v128.xor i32x4.extract_lane 0 call $log)"#, "240" },
    test_v128_not => { r#"(func (export "_start")
        v128.const i32x4 0 0 0 0 v128.not i32x4.extract_lane 0 call $log)"#, "-1" },
    test_v128_andnot => { r#"(func (export "_start")
        v128.const i32x4 0xFF 0 0 0 v128.const i32x4 0x0F 0 0 0
        v128.andnot i32x4.extract_lane 0 call $log)"#, "240" },
    test_v128_bitselect => { r#"(func (export "_start")
        v128.const i32x4 0xAAAA 0 0 0
        v128.const i32x4 0x5555 0 0 0
        v128.const i32x4 0xFF00 0 0 0
        v128.bitselect i32x4.extract_lane 0 call $log)"#, "43605" },
    test_v128_and_across_lanes => { r#"(func (export "_start")
        v128.const i16x8 0xFFFF 0xFFFF 0 0 0 0 0 0
        v128.const i16x8 0x00FF 0xFF00 0 0 0 0 0 0
        v128.and i16x8.extract_lane_u 1 call $log)"#, "65280" },
    test_v128_or_lane_high => { r#"(func (export "_start")
        v128.const i32x4 0 0 0 0x1 v128.const i32x4 0 0 0 0x2
        v128.or i32x4.extract_lane 3 call $log)"#, "3" },
    test_v128_xor_self_is_zero => { r#"(func (export "_start")
        v128.const i32x4 0x12345678 0 0 0
        v128.const i32x4 0x12345678 0 0 0
        v128.xor i32x4.extract_lane 0 call $log)"#, "0" },
}
