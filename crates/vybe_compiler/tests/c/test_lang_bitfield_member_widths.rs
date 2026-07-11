//! Bitfield member widths — 1-bit, 3-bit, signed fields, and masked observable values.

c_run_cases! {
    bitfield_single_one_bit_set => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; };",
        body: "struct F f; f.a = 1; printf(\"%u\\n\", f.a); return 0;",
        expect: ["1"]
    },
    bitfield_single_one_bit_clear => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; };",
        body: "struct F f; f.a = 0; printf(\"%u\\n\", f.a); return 0;",
        expect: ["0"]
    },
    bitfield_one_bit_wraps_to_zero => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; };",
        body: "struct F f; f.a = 2; printf(\"%u\\n\", f.a); return 0;",
        expect: ["0"]
    },
    bitfield_three_bit_stores_value => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 5; printf(\"%u\\n\", f.a); return 0;",
        expect: ["5"]
    },
    bitfield_three_bit_masks_eight_to_zero => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 8; printf(\"%u\\n\", f.a); return 0;",
        expect: ["0"]
    },
    bitfield_three_bit_masks_seven => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 7; printf(\"%u\\n\", f.a); return 0;",
        expect: ["7"]
    },
    bitfield_two_one_bit_fields_independent => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; unsigned b : 1; };",
        body: "struct F f; f.a = 1; f.b = 0; printf(\"%u %u\\n\", f.a, f.b); return 0;",
        expect: ["1 0"]
    },
    bitfield_three_bit_and_one_bit_together => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned lo : 3; unsigned hi : 1; };",
        body: "struct F f; f.lo = 4; f.hi = 1; printf(\"%u %u\\n\", f.lo, f.hi); return 0;",
        expect: ["4 1"]
    },
    bitfield_signed_four_bit_negative_one => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 4; };",
        body: "struct F f; f.s = -1; printf(\"%d\\n\", f.s); return 0;",
        expect: ["-1"]
    },
    bitfield_signed_four_bit_positive_three => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 4; };",
        body: "struct F f; f.s = 3; printf(\"%d\\n\", f.s); return 0;",
        expect: ["3"]
    },
    bitfield_signed_three_bit_negative_three => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 3; };",
        body: "struct F f; f.s = -3; printf(\"%d\\n\", f.s); return 0;",
        expect: ["-3"]
    },
    bitfield_signed_three_bit_positive_two => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 3; };",
        body: "struct F f; f.s = 2; printf(\"%d\\n\", f.s); return 0;",
        expect: ["2"]
    },
    bitfield_initializer_one_and_three_bits => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; unsigned b : 3; };",
        body: "struct F f = {1, 6}; printf(\"%u %u\\n\", f.a, f.b); return 0;",
        expect: ["1 6"]
    },
    bitfield_post_increment_one_bit => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; };",
        body: "struct F f; f.a = 0; f.a++; printf(\"%u\\n\", f.a); return 0;",
        expect: ["1"]
    },
    bitfield_post_increment_wraps_one_bit => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; };",
        body: "struct F f; f.a = 1; f.a++; printf(\"%u\\n\", f.a); return 0;",
        expect: ["0"]
    },
    bitfield_pre_increment_three_bit => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 6; ++f.a; printf(\"%u\\n\", f.a); return 0;",
        expect: ["7"]
    },
    bitfield_pre_increment_wraps_three_bit => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 7; ++f.a; printf(\"%u\\n\", f.a); return 0;",
        expect: ["0"]
    },
    bitfield_compound_add_assign_three_bit => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 2; f.a += 3; printf(\"%u\\n\", f.a); return 0;",
        expect: ["5"]
    },
    bitfield_compound_add_overflow_mask => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 5; f.a += 4; printf(\"%u\\n\", f.a); return 0;",
        expect: ["1"]
    },
    bitfield_and_assign_masks_width => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 7; f.a &= 3; printf(\"%u\\n\", f.a); return 0;",
        expect: ["3"]
    },
    bitfield_or_assign_within_width => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 1; f.a |= 4; printf(\"%u\\n\", f.a); return 0;",
        expect: ["5"]
    },
    bitfield_xor_assign_toggle_bits => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 5; f.a ^= 7; printf(\"%u\\n\", f.a); return 0;",
        expect: ["2"]
    },
    bitfield_with_regular_int_field => {
        includes: ["<stdio.h>"],
        decls: "struct F { int id; unsigned a : 1; unsigned b : 3; };",
        body: "struct F f; f.id = 9; f.a = 1; f.b = 4; printf(\"%d %u %u\\n\", f.id, f.a, f.b); return 0;",
        expect: ["9 1 4"]
    },
    bitfield_read_after_overwrite => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 7; f.a = 2; printf(\"%u\\n\", f.a); return 0;",
        expect: ["2"]
    },
    bitfield_two_three_bit_fields => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned x : 3; unsigned y : 3; };",
        body: "struct F f; f.x = 3; f.y = 6; printf(\"%u %u\\n\", f.x, f.y); return 0;",
        expect: ["3 6"]
    },
    bitfield_four_bit_unsigned_max => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 4; };",
        body: "struct F f; f.a = 15; printf(\"%u\\n\", f.a); return 0;",
        expect: ["15"]
    },
    bitfield_four_bit_overflow_sixteen => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 4; };",
        body: "struct F f; f.a = 16; printf(\"%u\\n\", f.a); return 0;",
        expect: ["0"]
    },
    bitfield_signed_four_bit_negative_four => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 4; };",
        body: "struct F f; f.s = -4; printf(\"%d\\n\", f.s); return 0;",
        expect: ["-4"]
    },
    bitfield_signed_five_bit_negative_one => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 5; };",
        body: "struct F f; f.s = -1; printf(\"%d\\n\", f.s); return 0;",
        expect: ["-1"]
    },
    bitfield_signed_assign_from_int_literal => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 4; };",
        body: "struct F f; f.s = -8; printf(\"%d\\n\", f.s); return 0;",
        expect: ["-8"]
    },
    bitfield_pointer_arrow_write_read => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; struct F *p = &f; p->a = 4; printf(\"%u\\n\", f.a); return 0;",
        expect: ["4"]
    },
    bitfield_in_struct_array => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F arr[2]; arr[0].a = 1; arr[1].a = 7; printf(\"%u %u\\n\", arr[0].a, arr[1].a); return 0;",
        expect: ["1 7"]
    },
    bitfield_global_storage => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; unsigned b : 3; }; struct F g = {1, 5};",
        body: "printf(\"%u %u\\n\", g.a, g.b); return 0;",
        expect: ["1 5"]
    },
    bitfield_compare_two_fields_equal => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; unsigned b : 3; };",
        body: "struct F f; f.a = 4; f.b = 4; printf(\"%d\\n\", f.a == f.b); return 0;",
        expect: ["1"]
    },
    bitfield_compare_less_than => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 2; printf(\"%d\\n\", f.a < 5); return 0;",
        expect: ["1"]
    },
    bitfield_ternary_select_field => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; unsigned b : 3; };",
        body: "struct F f; f.a = 0; f.b = 6; printf(\"%u\\n\", f.a ? 1 : f.b); return 0;",
        expect: ["6"]
    },
    bitfield_switch_on_three_bit_value => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 3; switch(f.a){case 3: printf(\"ok\\n\"); break; default: printf(\"no\\n\");} return 0;",
        expect: ["ok"]
    },
    bitfield_add_to_regular_field_unaffected => {
        includes: ["<stdio.h>"],
        decls: "struct F { int n; unsigned a : 3; };",
        body: "struct F f; f.n = 10; f.a = 3; f.n += 5; printf(\"%d %u\\n\", f.n, f.a); return 0;",
        expect: ["15 3"]
    },
    bitfield_left_shift_assign_within_width => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 1; f.a <<= 2; printf(\"%u\\n\", f.a); return 0;",
        expect: ["4"]
    },
    bitfield_right_shift_assign => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 6; f.a >>= 1; printf(\"%u\\n\", f.a); return 0;",
        expect: ["3"]
    },
    bitfield_modulo_assign => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 7; f.a %= 4; printf(\"%u\\n\", f.a); return 0;",
        expect: ["3"]
    },
    bitfield_multiply_assign_masks => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 3; f.a *= 3; printf(\"%u\\n\", f.a); return 0;",
        expect: ["1"]
    },
    bitfield_negation_signed_four_bit => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 4; };",
        body: "struct F f; f.s = 3; printf(\"%d\\n\", -f.s); return 0;",
        expect: ["-3"]
    },
    bitfield_three_one_bit_flags_pattern => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; unsigned b : 1; unsigned c : 1; };",
        body: "struct F f = {1, 0, 1}; printf(\"%u %u %u\\n\", f.a, f.b, f.c); return 0;",
        expect: ["1 0 1"]
    },
    bitfield_assign_from_expression => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; };",
        body: "struct F f; f.a = 1 + 2; printf(\"%u\\n\", f.a); return 0;",
        expect: ["3"]
    },
    bitfield_read_in_arithmetic => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 3; unsigned b : 3; };",
        body: "struct F f; f.a = 2; f.b = 3; printf(\"%u\\n\", f.a + f.b); return 0;",
        expect: ["5"]
    },
    bitfield_signed_three_bit_zero => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 3; };",
        body: "struct F f; f.s = 0; printf(\"%d\\n\", f.s); return 0;",
        expect: ["0"]
    },
    bitfield_unsigned_two_bit_max_three => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 2; };",
        body: "struct F f; f.a = 3; printf(\"%u\\n\", f.a); return 0;",
        expect: ["3"]
    },
    bitfield_unsigned_two_bit_wrap_four => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 2; };",
        body: "struct F f; f.a = 4; printf(\"%u\\n\", f.a); return 0;",
        expect: ["0"]
    },
    bitfield_nested_struct_container => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { unsigned a : 3; }; struct Outer { struct Inner in; };",
        body: "struct Outer o; o.in.a = 5; printf(\"%u\\n\", o.in.a); return 0;",
        expect: ["5"]
    },
    bitfield_copy_struct_preserves_fields => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a : 1; unsigned b : 3; };",
        body: "struct F a; a.a = 1; a.b = 6; struct F c = a; printf(\"%u %u\\n\", c.a, c.b); return 0;",
        expect: ["1 6"]
    },
    bitfield_decrement_signed_four_bit => {
        includes: ["<stdio.h>"],
        decls: "struct F { signed s : 4; };",
        body: "struct F f; f.s = 2; f.s--; printf(\"%d\\n\", f.s); return 0;",
        expect: ["1"]
    },
}
