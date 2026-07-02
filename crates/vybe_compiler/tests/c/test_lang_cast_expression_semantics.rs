//! Explicit cast expressions that change representation or observable values.


c_run_cases! {
    int_ninety_seven_to_char_prints_a => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%c\\n\", (char)97); return 0;",
        expect: ["a"]
    },
    int_sixty_five_to_char_prints_a => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%c\\n\", (char)65); return 0;",
        expect: ["A"]
    },
    int_zero_to_char_code => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(char)0); return 0;",
        expect: ["0"]
    },
    int_two_fifty_six_to_char_truncates => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(char)256); return 0;",
        expect: ["0"]
    },
    int_two_fifty_seven_to_char_truncates => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(char)257); return 0;",
        expect: ["1"]
    },
    int_three_hundred_to_char_truncates => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(char)300); return 0;",
        expect: ["44"]
    },
    int_negative_one_to_char_truncates => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(char)-1); return 0;",
        expect: ["-1"]
    },
    int_to_char_in_addition_expression => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%c\\n\", (char)(65 + 2)); return 0;",
        expect: ["C"]
    },
    int_to_unsigned_char_wraps => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", (unsigned char)256); return 0;",
        expect: ["0"]
    },
    int_to_signed_char_preserves_small_value => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (signed char)120); return 0;",
        expect: ["120"]
    },
    double_nine_point_nine_nine_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)9.99); return 0;",
        expect: ["9"]
    },
    double_two_point_three_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)2.3); return 0;",
        expect: ["2"]
    },
    double_negative_two_point_nine_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(-2.9)); return 0;",
        expect: ["-2"]
    },
    double_zero_point_nine_nine_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)0.99); return 0;",
        expect: ["0"]
    },
    double_one_hundred_point_zero_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)100.0); return 0;",
        expect: ["100"]
    },
    float_seven_point_eight_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)7.8f); return 0;",
        expect: ["7"]
    },
    double_negative_zero_point_one_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(-0.1)); return 0;",
        expect: ["0"]
    },
    double_large_fractional_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)999.999); return 0;",
        expect: ["999"]
    },
    cast_double_sum_before_truncation => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(4.2 + 3.8)); return 0;",
        expect: ["8"]
    },
    cast_float_product_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(3.5f * 2.0f)); return 0;",
        expect: ["7"]
    },
    void_pointer_to_int_pointer_deref => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x = 19; void *vp = &x; printf(\"%d\\n\", *(int *)vp); return 0;",
        expect: ["19"]
    },
    char_pointer_from_int_array_cast => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int arr[2] = {5, 6}; char *cp = (char *)arr; printf(\"%d\\n\", *(int *)cp); return 0;",
        expect: ["5"]
    },
    int_pointer_to_char_pointer_back => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n = 31; int *ip = &n; char *cp = (char *)ip; printf(\"%d\\n\", *(int *)cp); return 0;",
        expect: ["31"]
    },
    null_pointer_to_int_zero => {
        includes: ["<stdio.h>", "<stdint.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(intptr_t)(void *)0); return 0;",
        expect: ["0"]
    },
    intptr_to_pointer_reads_value => {
        includes: ["<stdio.h>", "<stdint.h>"],
        decls: "",
        body: "int val = 55; intptr_t ip = (intptr_t)&val; printf(\"%d\\n\", *(int *)ip); return 0;",
        expect: ["55"]
    },
    uintptr_roundtrip_preserves_deref => {
        includes: ["<stdio.h>", "<stdint.h>"],
        decls: "",
        body: "int val = 77; uintptr_t up = (uintptr_t)&val; int *back = (int *)up; printf(\"%d\\n\", *back); return 0;",
        expect: ["77"]
    },
    const_int_pointer_cast_mutation => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "const int c = 12; int *p = (int *)&c; *p = 13; printf(\"%d\\n\", c); return 0;",
        expect: ["13"]
    },
    double_pointer_to_void_and_back => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "double d = 3.5; void *vp = (void *)&d; printf(\"%.1f\\n\", *(double *)vp); return 0;",
        expect: ["3.5"]
    },
    short_to_int_to_char_chain => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%c\\n\", (char)(int)(short)66); return 0;",
        expect: ["B"]
    },
    int_to_short_truncates_bits => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (short)70000); return 0;",
        expect: ["4464"]
    },
    unsigned_int_to_signed_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(unsigned int)2147483648u); return 0;",
        expect: ["-2147483648"]
    },
    signed_int_to_unsigned_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", (unsigned int)-2); return 0;",
        expect: ["4294967294"]
    },
    float_to_int_truncates_in_assignment => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n = (int)8.7f; printf(\"%d\\n\", n); return 0;",
        expect: ["8"]
    },
    double_to_char_via_int_truncation => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%c\\n\", (char)(int)66.9); return 0;",
        expect: ["B"]
    },
    pointer_to_array_element_cast => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int arr[3] = {10, 20, 30}; int *mid = (int *)&arr[1]; printf(\"%d\\n\", *mid); return 0;",
        expect: ["20"]
    },
    cast_pointer_difference_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int arr[4] = {0}; printf(\"%d\\n\", (int)(&arr[3] - &arr[0])); return 0;",
        expect: ["3"]
    },
    cast_bool_like_int_to_char => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (char)(5 > 3)); return 0;",
        expect: ["1"]
    },
    nested_int_to_char_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(char)(int)(char)90); return 0;",
        expect: ["90"]
    },
    double_neg_nine_point_one_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(-9.1)); return 0;",
        expect: ["-9"]
    },
    long_double_to_int_truncates => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)11.95L); return 0;",
        expect: ["11"]
    },
    int_to_float_back_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(float)47); return 0;",
        expect: ["47"]
    },
    char_to_int_promotion_in_cast => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)'Z'); return 0;",
        expect: ["90"]
    },
    void_pointer_from_int_pointer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x = 63; int *ip = &x; void *vp = (void *)ip; printf(\"%d\\n\", *(int *)vp); return 0;",
        expect: ["63"]
    },
    cast_in_ternary_branch => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 1 ? (int)4.9 : (int)1.1); return 0;",
        expect: ["4"]
    },
    cast_in_comma_expression => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", ((int)1.2, (int)3.8)); return 0;",
        expect: ["3"]
    },
    int_pointer_cast_from_null_compares_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int *p = (int *)0; printf(\"%d\\n\", p == 0); return 0;",
        expect: ["1"]
    },
    double_to_unsigned_int_truncates => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", (unsigned int)12.9); return 0;",
        expect: ["12"]
    },
    int_to_long_preserves_in_printf => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%ld\\n\", (long)123456789); return 0;",
        expect: ["123456789"]
    },
    array_decay_cast_to_pointer_deref => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int arr[2] = {88, 99}; int *p = (int *)arr; printf(\"%d\\n\", p[1]); return 0;",
        expect: ["99"]
    },
    cast_float_literal_in_comparison => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)5.0f == 5); return 0;",
        expect: ["1"]
    },
}

c_compile_cases! {
    function_pointer_to_void_pointer => {
        includes: ["<stdio.h>"],
        decls: "int id(int x) { return x; }",
        body: "int (*fp)(int) = id; void *vp = (void *)fp; return vp != 0;"
    },
    incompatible_struct_pointer_cast => {
        includes: ["<stdio.h>"],
        decls: "struct A { int x; }; struct B { int y; };",
        body: "struct A a = {1}; struct B *bp = (struct B *)&a; return bp->y;"
    },
}
