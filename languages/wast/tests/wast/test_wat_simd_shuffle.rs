//! SIMD byte shuffle and lane conversions between shapes.
use crate::wat_exec;

wat_exec! {
    test_i8x16_shuffle_identity_low => { r#"(func (export "_start")
        v128.const i8x16 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25
        v128.const i8x16 100 101 102 103 104 105 106 107 108 109 110 111 112 113 114 115
        i8x16.shuffle 0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15
        i8x16.extract_lane_u 3 call $log)"#, "13" },
    test_i8x16_shuffle_from_second_vector => { r#"(func (export "_start")
        v128.const i8x16 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25
        v128.const i8x16 100 101 102 103 104 105 106 107 108 109 110 111 112 113 114 115
        i8x16.shuffle 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31
        i8x16.extract_lane_u 0 call $log)"#, "100" },
    test_i8x16_shuffle_interleave => { r#"(func (export "_start")
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        v128.const i8x16 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9 9
        i8x16.shuffle 0 16 1 17 2 18 3 19 4 20 5 21 6 22 7 23
        i8x16.extract_lane_u 1 call $log)"#, "9" },
    test_i8x16_shuffle_reverse => { r#"(func (export "_start")
        v128.const i8x16 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16
        v128.const i8x16 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.shuffle 15 14 13 12 11 10 9 8 7 6 5 4 3 2 1 0
        i8x16.extract_lane_u 0 call $log)"#, "16" },
    test_i8x16_swizzle_by_index => { r#"(func (export "_start")
        v128.const i8x16 10 20 30 40 50 60 70 80 90 100 110 120 130 140 150 160
        v128.const i8x16 5 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i8x16.swizzle i8x16.extract_lane_u 0 call $log)"#, "60" },
    test_i16x8_widen_extract => { r#"(func (export "_start")
        v128.const i8x16 200 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0
        i16x8.extend_low_i8x16_u i16x8.extract_lane_u 0 call $log)"#, "200" },
}
