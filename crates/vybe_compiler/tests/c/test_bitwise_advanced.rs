use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    bit_mask_can_clear_low_bit => { body: "printf(\"%d\\n\", 7 & ~1); return 0;", expect: ["6"] },
    bit_mask_can_set_low_bit => { body: "printf(\"%d\\n\", 6 | 1); return 0;", expect: ["7"] },
    bit_mask_can_toggle_middle_bit => { body: "printf(\"%d\\n\", 7 ^ 2); return 0;", expect: ["5"] },
    left_shift_by_zero_keeps_value => { body: "printf(\"%d\\n\", 5 << 0); return 0;", expect: ["5"] },
    right_shift_by_zero_keeps_value => { body: "printf(\"%d\\n\", 5 >> 0); return 0;", expect: ["5"] },
    left_shift_can_double_positive_integer => { body: "printf(\"%d\\n\", 9 << 1); return 0;", expect: ["18"] },
    right_shift_can_halve_even_integer => { body: "printf(\"%d\\n\", 18 >> 1); return 0;", expect: ["9"] },
    bitwise_not_of_negative_one_is_zero => { body: "int x = -1; printf(\"%d\\n\", ~x); return 0;", expect: ["0"] },
    bitwise_and_has_higher_precedence_than_xor => { body: "printf(\"%d\\n\", 7 ^ 3 & 1); return 0;", expect: ["6"] },
    bitwise_xor_has_higher_precedence_than_or => { body: "printf(\"%d\\n\", 4 | 3 ^ 1); return 0;", expect: ["6"] },
    shift_then_mask_can_extract_field => { body: "int x = 0b11010; printf(\"%d\\n\", (x >> 1) & 0b11); return 0;", expect: ["1"] },
    mask_can_test_bit_presence => { body: "int x = 0b1010; printf(\"%d\\n\", (x & 0b1000) != 0); return 0;", expect: ["1"] },
    mask_can_test_bit_absence => { body: "int x = 0b1010; printf(\"%d\\n\", (x & 0b0100) != 0); return 0;", expect: ["0"] },
    xor_of_value_with_itself_is_zero => { body: "printf(\"%d\\n\", 123 ^ 123); return 0;", expect: ["0"] },
    and_of_value_with_itself_is_identity => { body: "printf(\"%d\\n\", 123 & 123); return 0;", expect: ["123"] },
    or_of_value_with_itself_is_identity => { body: "printf(\"%d\\n\", 123 | 123); return 0;", expect: ["123"] },
    shifting_then_unshifting_even_value_restores_original => { body: "printf(\"%d\\n\", (24 << 1) >> 1); return 0;", expect: ["24"] },
    bitwise_ops_can_drive_condition => { body: "if ((6 & 2) != 0) puts(\"set\"); else puts(\"clear\"); return 0;", expect: ["set"] },
    low_nibble_mask_can_extract_bottom_four_bits => { body: "printf(\"%d\\n\", 0xAB & 0x0F); return 0;", expect: ["11"] },
    high_nibble_mask_can_extract_top_four_bits => { body: "printf(\"%d\\n\", (0xAB & 0xF0) >> 4); return 0;", expect: ["10"] },
    xor_can_toggle_sparse_pattern => { body: "printf(\"%d\\n\", 0b1111 ^ 0b0101); return 0;", expect: ["10"] },
    and_can_zero_non_matching_bits => { body: "printf(\"%d\\n\", 0b1111 & 0b0101); return 0;", expect: ["5"] },
    or_can_merge_non_overlapping_patterns => { body: "printf(\"%d\\n\", 0b1000 | 0b0011); return 0;", expect: ["11"] },
    shift_expression_can_feed_array_index => { body: "int arr[4] = {1, 2, 3, 4}; printf(\"%d\\n\", arr[1 << 1]); return 0;", expect: ["3"] },
    bitwise_not_of_zero_is_minus_one => { body: "printf(\"%d\\n\", ~0); return 0;", expect: ["-1"] },
    left_shift_on_negative_one_yields_negative_two => { body: "printf(\"%d\\n\", -1 << 1); return 0;", expect: ["-2"] },
    right_shift_on_large_positive_discards_low_bits => { body: "printf(\"%d\\n\", 255 >> 4); return 0;", expect: ["15"] },
    boolean_comparison_of_bitwise_results_can_be_true => { body: "printf(\"%d\\n\", ((5 & 1) == 1) && ((5 >> 2) == 1)); return 0;", expect: ["1"] },
    multiple_bitwise_ops_can_chain_left_to_right => { body: "printf(\"%d\\n\", (12 ^ 10) | 1); return 0;", expect: ["7"] },
    bitwise_expression_can_be_cast_to_char => { body: "printf(\"%c\\n\", (char)(0x40 | 0x01)); return 0;", expect: ["A"] },
    bitwise_and_with_zero_is_zero => { body: "printf(\"%d\\n\", 123 & 0); return 0;", expect: ["0"] },
    bitwise_or_with_zero_is_identity => { body: "printf(\"%d\\n\", 123 | 0); return 0;", expect: ["123"] },
    bitwise_xor_with_zero_is_identity => { body: "printf(\"%d\\n\", 123 ^ 0); return 0;", expect: ["123"] },
    shifted_mask_can_select_single_middle_bit => { body: "int x = 0b10100; printf(\"%d\\n\", (x & (1 << 2)) != 0); return 0;", expect: ["1"] },
    combined_shifts_can_restore_power_of_two => { body: "printf(\"%d\\n\", (1 << 5) >> 5); return 0;", expect: ["1"] },
    xor_can_toggle_flag_to_zero => { body: "int flags = 0b0010; flags ^= 0b0010; printf(\"%d\\n\", flags); return 0;", expect: ["0"] },
    and_then_or_sequence_can_replace_pattern => { body: "int x = 0b1111; x = (x & ~0b0110) | 0b0010; printf(\"%d\\n\", x); return 0;", expect: ["11"] },
    shift_result_can_participate_in_addition => { body: "printf(\"%d\\n\", (3 << 2) + 1); return 0;", expect: ["13"] },
    bitwise_expression_can_feed_ternary_branch => { body: "puts((8 & 8) ? \"set\" : \"clear\"); return 0;", expect: ["set"] },
    low_bit_test_of_even_number_is_zero => { body: "printf(\"%d\\n\", 8 & 1); return 0;", expect: ["0"] }
}
