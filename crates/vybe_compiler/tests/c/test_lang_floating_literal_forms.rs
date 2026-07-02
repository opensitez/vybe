//! Floating literal lexical forms: decimal, scientific, suffixes, hex floats.


c_run_cases! {
    one_point_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 1.0); return 0;",
        expect: ["1.0"]
    },
    dot_five => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", .5); return 0;",
        expect: ["0.5"]
    },
    one_e_ten => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 1e10); return 0;",
        expect: ["10000000000"]
    },
    one_f_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "float f = 1.f; printf(\"%.0f\\n\", (double)f); return 0;",
        expect: ["1"]
    },
    one_point_zero_l_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", (double)1.0L); return 0;",
        expect: ["1.0"]
    },
    dot_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", .0); return 0;",
        expect: ["0.0"]
    },
    zero_point_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 0.0); return 0;",
        expect: ["0.0"]
    },
    two_point_five => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 2.5); return 0;",
        expect: ["2.5"]
    },
    trailing_dot_one => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 1.); return 0;",
        expect: ["1"]
    },
    leading_dot_two_five => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.2f\\n\", .25); return 0;",
        expect: ["0.25"]
    },
    one_e_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 1e0); return 0;",
        expect: ["1"]
    },
    two_e_plus_three => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 2E+3); return 0;",
        expect: ["2000"]
    },
    five_e_minus_one => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 5e-1); return 0;",
        expect: ["0.5"]
    },
    three_dot_one_four_f => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "float f = 3.14f; printf(\"%.2f\\n\", (double)f); return 0;",
        expect: ["3.14"]
    },
    two_dot_five_f => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "float f = 2.5F; printf(\"%.1f\\n\", (double)f); return 0;",
        expect: ["2.5"]
    },
    two_dot_zero_l => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "long double ld = 2.0L; printf(\"%.1f\\n\", (double)ld); return 0;",
        expect: ["2.0"]
    },
    scientific_no_decimal_part => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 1e2); return 0;",
        expect: ["100"]
    },
    one_dot_zero_e_one => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 1.0e1); return 0;",
        expect: ["10"]
    },
    negative_one_point_five => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", -1.5); return 0;",
        expect: ["-1.5"]
    },
    unary_plus_two_point_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", +2.0); return 0;",
        expect: ["2.0"]
    },
    zero_point_five => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 0.5); return 0;",
        expect: ["0.5"]
    },
    ten_point_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 10.0); return 0;",
        expect: ["10.0"]
    },
    nine_dot_nine_nine_e_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.2f\\n\", 9.99e0); return 0;",
        expect: ["9.99"]
    },
    one_dot_two_three_e_minus_two => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.4f\\n\", 1.23e-2); return 0;",
        expect: ["0.0123"]
    },
    one_dot_two_three_e_plus_two => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 1.23e+2); return 0;",
        expect: ["123"]
    },
    leading_zeros_before_decimal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 00001.0); return 0;",
        expect: ["1.0"]
    },
    trailing_fractional_zeros => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 1.000); return 0;",
        expect: ["1"]
    },
    zero_dot_alone => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 0.); return 0;",
        expect: ["0.0"]
    },
    one_e_plus_ten => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 1e+10); return 0;",
        expect: ["10000000000"]
    },
    one_e_minus_four => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.4f\\n\", 1E-4); return 0;",
        expect: ["0.0001"]
    },
    float_literal_addition => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 1.0 + 2.0); return 0;",
        expect: ["3.0"]
    },
    float_literal_subtraction => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 5.0 - 2.5); return 0;",
        expect: ["2.5"]
    },
    float_literal_multiplication => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 2.0 * 3.0); return 0;",
        expect: ["6.0"]
    },
    float_literal_division => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 6.0 / 2.0); return 0;",
        expect: ["3.0"]
    },
    float_literals_compare_equal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 1.0 == 1.0); return 0;",
        expect: ["1"]
    },
    cast_f_suffix_to_double => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.2f\\n\", (double)1.f); return 0;",
        expect: ["1.00"]
    },
    long_double_literal_to_double => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.2f\\n\", (double)3.0L); return 0;",
        expect: ["3.00"]
    },
    four_dot_zero_f_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "float f = 4.0f; printf(\"%.0f\\n\", (double)f); return 0;",
        expect: ["4"]
    },
    eight_dot_zero_l_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", (double)8.0L); return 0;",
        expect: ["8"]
    },
    six_dot_zero_e_one => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 6.0e1); return 0;",
        expect: ["60"]
    },
    seven_e_two => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 7e2); return 0;",
        expect: ["700"]
    },
    dot_one_two_five => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.3f\\n\", .125); return 0;",
        expect: ["0.125"]
    },
    one_dot_zero_f_in_expression => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", (double)(1.f + 2.f)); return 0;",
        expect: ["3"]
    },
    double_literal_in_modulo_context => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)(10.0 / 3.0)); return 0;",
        expect: ["3"]
    },
    hexfloat_one_dot_eight_p_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 0x1.8p0); return 0;",
        expect: ["1.5"]
    },
    hexfloat_one_p_one => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 0x1.0p1); return 0;",
        expect: ["2"]
    },
    hexfloat_one_p_minus_one => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.1f\\n\", 0x1.0p-1); return 0;",
        expect: ["0.5"]
    },
    hexfloat_a_p_two => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 0xa.p2); return 0;",
        expect: ["40"]
    },
    hexfloat_uppercase_prefix_and_exponent => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%.0f\\n\", 0X1.0P0); return 0;",
        expect: ["1"]
    },
}

c_compile_cases! {
    hexfloat_f_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "float f = 0x1.0p0f; return (int)f;"
    },
    hexfloat_l_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "long double ld = 0x1.0p0L; return (int)ld;"
    },
    hexfloat_large_exponent => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "double d = 0x1.0p10; return (int)d;"
    },
    hexfloat_fractional_mantissa => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "double d = 0x1.FFp0; return (int)d;"
    },
    hexfloat_negative_exponent => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "double d = 0x1.0p-4; return (int)(d * 16.0);"
    },
}
