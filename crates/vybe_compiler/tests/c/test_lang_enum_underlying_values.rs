//! Enum underlying values, gaps, comparisons, switch, and mixed enum/int use.

use crate::helpers::*;

c_run_cases! {
    enum_explicit_start_then_auto_increment => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 5, B, C };",
        body: "printf(\"%d %d %d\\n\", A, B, C); return 0;",
        expect: ["5 6 7"]
    },
    enum_large_gap_between_constants => {
        includes: ["<stdio.h>"],
        decls: "enum Code { OK = 200, CREATED = 201, ACCEPTED = 202 };",
        body: "printf(\"%d %d\\n\", OK, ACCEPTED); return 0;",
        expect: ["200 202"]
    },
    enum_sparse_values_no_auto_fill => {
        includes: ["<stdio.h>"],
        decls: "enum E { X = 10, Y = 20, Z = 30 };",
        body: "printf(\"%d %d %d\\n\", X, Y, Z); return 0;",
        expect: ["10 20 30"]
    },
    enum_mixed_explicit_and_implicit_after_gap => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 1, B = 10, C, D = 20, E };",
        body: "printf(\"%d %d %d %d %d\\n\", A, B, C, D, E); return 0;",
        expect: ["1 10 11 20 21"]
    },
    enum_compare_same_constant_true => {
        includes: ["<stdio.h>"],
        decls: "enum S { OFF, ON };",
        body: "enum S s = ON; printf(\"%d\\n\", s == ON); return 0;",
        expect: ["1"]
    },
    enum_compare_different_constants_false => {
        includes: ["<stdio.h>"],
        decls: "enum S { OFF, ON };",
        body: "enum S s = OFF; printf(\"%d\\n\", s == ON); return 0;",
        expect: ["0"]
    },
    enum_compare_less_than_by_value => {
        includes: ["<stdio.h>"],
        decls: "enum R { LOW = 1, HIGH = 9 };",
        body: "enum R r = LOW; printf(\"%d\\n\", r < HIGH); return 0;",
        expect: ["1"]
    },
    enum_switch_hits_matching_case => {
        includes: ["<stdio.h>"],
        decls: "enum Color { RED, GREEN, BLUE };",
        body: "enum Color c = GREEN; switch(c){case RED: printf(\"r\\n\"); break; case GREEN: printf(\"g\\n\"); break; default: printf(\"x\\n\");} return 0;",
        expect: ["g"]
    },
    enum_switch_falls_to_default => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 1, B = 2 };",
        body: "enum E e = (enum E)99; switch(e){case A: printf(\"a\\n\"); break; case B: printf(\"b\\n\"); break; default: printf(\"d\\n\");} return 0;",
        expect: ["d"]
    },
    enum_switch_with_explicit_value_labels => {
        includes: ["<stdio.h>"],
        decls: "enum Http { OK = 200, NOT_FOUND = 404 };",
        body: "enum Http h = NOT_FOUND; switch(h){case OK: printf(\"ok\\n\"); break; case NOT_FOUND: printf(\"nf\\n\"); break;} return 0;",
        expect: ["nf"]
    },
    enum_mixed_with_int_addition => {
        includes: ["<stdio.h>"],
        decls: "enum N { TWO = 2, THREE = 3 };",
        body: "printf(\"%d\\n\", TWO + THREE); return 0;",
        expect: ["5"]
    },
    enum_variable_plus_int_literal => {
        includes: ["<stdio.h>"],
        decls: "enum N { BASE = 10 };",
        body: "enum N n = BASE; printf(\"%d\\n\", n + 5); return 0;",
        expect: ["15"]
    },
    enum_assign_from_compatible_int => {
        includes: ["<stdio.h>"],
        decls: "enum E { A, B, C };",
        body: "enum E e = (enum E)2; printf(\"%d\\n\", e); return 0;",
        expect: ["2"]
    },
    enum_increment_wraps_as_int => {
        includes: ["<stdio.h>"],
        decls: "enum C { ZERO, ONE, TWO };",
        body: "enum C c = ONE; c = (enum C)(c + 1); printf(\"%d\\n\", c); return 0;",
        expect: ["2"]
    },
    enum_decrement_by_subtraction => {
        includes: ["<stdio.h>"],
        decls: "enum C { ZERO, ONE, TWO };",
        body: "enum C c = TWO; c = (enum C)(c - 1); printf(\"%d\\n\", c); return 0;",
        expect: ["1"]
    },
    enum_negative_explicit_values => {
        includes: ["<stdio.h>"],
        decls: "enum S { NEG = -3, ZERO = 0, POS = 3 };",
        body: "printf(\"%d %d\\n\", NEG, POS); return 0;",
        expect: ["-3 3"]
    },
    enum_negative_then_auto_increment => {
        includes: ["<stdio.h>"],
        decls: "enum S { A = -2, B, C };",
        body: "printf(\"%d %d %d\\n\", A, B, C); return 0;",
        expect: ["-2 -1 0"]
    },
    enum_in_struct_field_switch => {
        includes: ["<stdio.h>"],
        decls: "enum Mode { IDLE, RUN }; struct Job { enum Mode mode; };",
        body: "struct Job j = {RUN}; switch(j.mode){case IDLE: printf(\"i\\n\"); break; case RUN: printf(\"r\\n\"); break;} return 0;",
        expect: ["r"]
    },
    enum_struct_field_compare_to_constant => {
        includes: ["<stdio.h>"],
        decls: "enum L { LOW, MID, HIGH }; struct S { enum L level; };",
        body: "struct S s = {HIGH}; printf(\"%d\\n\", s.level == HIGH); return 0;",
        expect: ["1"]
    },
    enum_array_indexed_by_enum_value => {
        includes: ["<stdio.h>"],
        decls: "enum I { I0, I1, I2 };",
        body: "int vals[3] = {10, 20, 30}; printf(\"%d\\n\", vals[I2]); return 0;",
        expect: ["30"]
    },
    enum_typedef_switch_case => {
        includes: ["<stdio.h>"],
        decls: "typedef enum { EAST, WEST } Dir;",
        body: "Dir d = WEST; switch(d){case EAST: printf(\"e\\n\"); break; case WEST: printf(\"w\\n\"); break;} return 0;",
        expect: ["w"]
    },
    enum_equality_with_int_literal => {
        includes: ["<stdio.h>"],
        decls: "enum E { VAL = 7 };",
        body: "enum E e = VAL; printf(\"%d\\n\", e == 7); return 0;",
        expect: ["1"]
    },
    enum_inequality_with_other_constant => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 1, B = 2 };",
        body: "printf(\"%d\\n\", A != B); return 0;",
        expect: ["1"]
    },
    enum_greater_equal_chain => {
        includes: ["<stdio.h>"],
        decls: "enum R { R1 = 1, R2 = 2, R3 = 3 };",
        body: "printf(\"%d\\n\", R3 >= R1); return 0;",
        expect: ["1"]
    },
    enum_switch_fallthrough_two_cases => {
        includes: ["<stdio.h>"],
        decls: "enum E { A, B, C };",
        body: "enum E e = A; switch(e){case A: case B: printf(\"ab\\n\"); break; case C: printf(\"c\\n\"); break;} return 0;",
        expect: ["ab"]
    },
    enum_multiply_constants => {
        includes: ["<stdio.h>"],
        decls: "enum N { TWO = 2, FOUR = 4 };",
        body: "printf(\"%d\\n\", TWO * FOUR); return 0;",
        expect: ["8"]
    },
    enum_subtract_constants => {
        includes: ["<stdio.h>"],
        decls: "enum N { TEN = 10, THREE = 3 };",
        body: "printf(\"%d\\n\", TEN - THREE); return 0;",
        expect: ["7"]
    },
    enum_modulo_constant_expression => {
        includes: ["<stdio.h>"],
        decls: "enum N { TEN = 10, THREE = 3 };",
        body: "printf(\"%d\\n\", TEN % THREE); return 0;",
        expect: ["1"]
    },
    enum_bitwise_or_of_constants => {
        includes: ["<stdio.h>"],
        decls: "enum F { A = 1, B = 2, C = 4 };",
        body: "printf(\"%d\\n\", A | C); return 0;",
        expect: ["5"]
    },
    enum_bitwise_and_of_constants => {
        includes: ["<stdio.h>"],
        decls: "enum F { A = 3, B = 5 };",
        body: "printf(\"%d\\n\", A & B); return 0;",
        expect: ["1"]
    },
    enum_ternary_with_enum_condition => {
        includes: ["<stdio.h>"],
        decls: "enum B { NO, YES };",
        body: "enum B b = YES; printf(\"%d\\n\", b ? 1 : 0); return 0;",
        expect: ["1"]
    },
    enum_return_from_function_matches => {
        includes: ["<stdio.h>"],
        decls: "enum E { X = 4, Y = 5 }; enum E pick(int n) { return n ? Y : X; }",
        body: "printf(\"%d\\n\", pick(1)); return 0;",
        expect: ["5"]
    },
    enum_param_compare_in_function => {
        includes: ["<stdio.h>"],
        decls: "enum E { A, B }; int is_b(enum E e) { return e == B; }",
        body: "printf(\"%d\\n\", is_b(B)); return 0;",
        expect: ["1"]
    },
    enum_reassign_to_different_constant => {
        includes: ["<stdio.h>"],
        decls: "enum S { OFF, ON };",
        body: "enum S s = OFF; s = ON; printf(\"%d\\n\", s); return 0;",
        expect: ["1"]
    },
    enum_global_initializer_explicit => {
        includes: ["<stdio.h>"],
        decls: "enum E { K = 42 }; enum E g = K;",
        body: "printf(\"%d\\n\", g); return 0;",
        expect: ["42"]
    },
    enum_in_for_loop_bound => {
        includes: ["<stdio.h>"],
        decls: "enum N { LEN = 3 };",
        body: "int s=0,i; for(i=0;i<LEN;i++) s+=i; printf(\"%d\\n\", s); return 0;",
        expect: ["3"]
    },
    enum_switch_nested_in_if => {
        includes: ["<stdio.h>"],
        decls: "enum E { A, B };",
        body: "enum E e = B; if(1){ switch(e){case A: printf(\"a\\n\"); break; case B: printf(\"b\\n\"); break;} } return 0;",
        expect: ["b"]
    },
    enum_compare_after_cast_from_int => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 0, B = 1 };",
        body: "enum E e = (enum E)1; printf(\"%d\\n\", e == B); return 0;",
        expect: ["1"]
    },
    enum_hundred_step_gap => {
        includes: ["<stdio.h>"],
        decls: "enum E { P = 100, Q = 200, R = 300 };",
        body: "printf(\"%d\\n\", Q - P); return 0;",
        expect: ["100"]
    },
    enum_zero_explicit_among_others => {
        includes: ["<stdio.h>"],
        decls: "enum E { Z = 0, O = 1, T = 2 };",
        body: "printf(\"%d %d\\n\", Z, T); return 0;",
        expect: ["0 2"]
    },
    enum_switch_first_constant => {
        includes: ["<stdio.h>"],
        decls: "enum E { FIRST, SECOND };",
        body: "enum E e = FIRST; switch(e){case FIRST: printf(\"1\\n\"); break; case SECOND: printf(\"2\\n\"); break;} return 0;",
        expect: ["1"]
    },
    enum_switch_last_constant => {
        includes: ["<stdio.h>"],
        decls: "enum E { FIRST, SECOND };",
        body: "enum E e = SECOND; switch(e){case FIRST: printf(\"1\\n\"); break; case SECOND: printf(\"2\\n\"); break;} return 0;",
        expect: ["2"]
    },
    enum_mixed_add_int_and_enum => {
        includes: ["<stdio.h>"],
        decls: "enum B { BASE = 5 };",
        body: "int x = BASE + 2; printf(\"%d\\n\", x); return 0;",
        expect: ["7"]
    },
    enum_pointer_to_enum_var => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 3 };",
        body: "enum E e = A; enum E *p = &e; *p = (enum E)4; printf(\"%d\\n\", e); return 0;",
        expect: ["4"]
    },
    enum_struct_pointer_update_field => {
        includes: ["<stdio.h>"],
        decls: "enum S { OFF, ON }; struct D { enum S s; };",
        body: "struct D d = {OFF}; struct D *p = &d; p->s = ON; printf(\"%d\\n\", d.s); return 0;",
        expect: ["1"]
    },
    enum_logical_and_in_condition => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 1, B = 2 };",
        body: "printf(\"%d\\n\", A && B); return 0;",
        expect: ["1"]
    },
    enum_unary_minus_constant => {
        includes: ["<stdio.h>"],
        decls: "enum E { P = 5 };",
        body: "printf(\"%d\\n\", -P); return 0;",
        expect: ["-5"]
    },
    enum_shift_left_constant => {
        includes: ["<stdio.h>"],
        decls: "enum E { N = 2 };",
        body: "printf(\"%d\\n\", N << 2); return 0;",
        expect: ["8"]
    },
    enum_shift_right_constant => {
        includes: ["<stdio.h>"],
        decls: "enum E { N = 8 };",
        body: "printf(\"%d\\n\", N >> 1); return 0;",
        expect: ["4"]
    },
    enum_chained_compare_less_equal => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 2, B = 2 };",
        body: "printf(\"%d\\n\", A <= B); return 0;",
        expect: ["1"]
    },
    enum_print_via_enum_variable => {
        includes: ["<stdio.h>"],
        decls: "enum E { V = 88 };",
        body: "enum E e = V; printf(\"%d\\n\", e); return 0;",
        expect: ["88"]
    },
    enum_switch_case_with_gap_value => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 1, B = 5, C = 6 };",
        body: "enum E e = B; switch(e){case A: printf(\"a\\n\"); break; case B: printf(\"b\\n\"); break; case C: printf(\"c\\n\"); break;} return 0;",
        expect: ["b"]
    },
}
