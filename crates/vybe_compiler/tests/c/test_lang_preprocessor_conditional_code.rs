//! Preprocessor conditional compilation — #if/#elif/#else with observable int output.

c_run_cases! {
    if_true_branch_selects_value => {
        includes: ["<stdio.h>"],
        decls: "#if 1\n#define VAL 11\n#else\n#define VAL 0\n#endif",
        body: "printf(\"%d\\n\", VAL); return 0;",
        expect: ["11"]
    },
    if_false_branch_skips_to_else => {
        includes: ["<stdio.h>"],
        decls: "#if 0\n#define VAL 1\n#else\n#define VAL 22\n#endif",
        body: "printf(\"%d\\n\", VAL); return 0;",
        expect: ["22"]
    },
    elif_second_branch_matches => {
        includes: ["<stdio.h>"],
        decls: "#define TIER 2\n#if TIER==1\n#define SCORE 10\n#elif TIER==2\n#define SCORE 20\n#else\n#define SCORE 0\n#endif",
        body: "printf(\"%d\\n\", SCORE); return 0;",
        expect: ["20"]
    },
    elif_third_branch_matches => {
        includes: ["<stdio.h>"],
        decls: "#define TIER 3\n#if TIER==1\n#define SCORE 10\n#elif TIER==2\n#define SCORE 20\n#elif TIER==3\n#define SCORE 30\n#else\n#define SCORE 0\n#endif",
        body: "printf(\"%d\\n\", SCORE); return 0;",
        expect: ["30"]
    },
    elif_chain_falls_through_to_else => {
        includes: ["<stdio.h>"],
        decls: "#define TIER 9\n#if TIER==1\n#define SCORE 10\n#elif TIER==2\n#define SCORE 20\n#else\n#define SCORE 99\n#endif",
        body: "printf(\"%d\\n\", SCORE); return 0;",
        expect: ["99"]
    },
    if_defined_macro_selects_one => {
        includes: ["<stdio.h>"],
        decls: "#define FEATURE\n#if defined(FEATURE)\n#define ON 1\n#else\n#define ON 0\n#endif",
        body: "printf(\"%d\\n\", ON); return 0;",
        expect: ["1"]
    },
    if_not_defined_selects_other => {
        includes: ["<stdio.h>"],
        decls: "#if !defined(MISSING)\n#define ON 2\n#else\n#define ON 0\n#endif",
        body: "printf(\"%d\\n\", ON); return 0;",
        expect: ["2"]
    },
    ifdef_defined_macro => {
        includes: ["<stdio.h>"],
        decls: "#define FLAG\n#ifdef FLAG\n#define V 3\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["3"]
    },
    ifndef_undefined_macro => {
        includes: ["<stdio.h>"],
        decls: "#ifndef GUARD\n#define V 4\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["4"]
    },
    if_arithmetic_comparison => {
        includes: ["<stdio.h>"],
        decls: "#define A 2\n#define B 3\n#if A+B==5\n#define V 5\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["5"]
    },
    if_less_than_comparison => {
        includes: ["<stdio.h>"],
        decls: "#define N 4\n#if N<10\n#define V 6\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["6"]
    },
    if_greater_equal_comparison => {
        includes: ["<stdio.h>"],
        decls: "#define N 10\n#if N>=10\n#define V 7\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["7"]
    },
    if_not_equal_comparison => {
        includes: ["<stdio.h>"],
        decls: "#define N 3\n#if N!=7\n#define V 8\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["8"]
    },
    nested_if_both_true => {
        includes: ["<stdio.h>"],
        decls: "#define A\n#define B\n#if defined(A)\n  #if defined(B)\n    #define V 9\n  #else\n    #define V 1\n  #endif\n#else\n  #define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["9"]
    },
    nested_if_inner_false => {
        includes: ["<stdio.h>"],
        decls: "#define A\n#if defined(A)\n  #if defined(Z)\n    #define V 1\n  #else\n    #define V 10\n  #endif\n#else\n  #define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["10"]
    },
    conditional_array_size_via_if => {
        includes: ["<stdio.h>"],
        decls: "#define BIG 1\n#if BIG\n#define N 4\n#else\n#define N 2\n#endif",
        body: "int a[N]={1,2,3,4}; printf(\"%d\\n\", a[N-1]); return 0;",
        expect: ["4"]
    },
    if_zero_is_false => {
        includes: ["<stdio.h>"],
        decls: "#if 0\n#define V 1\n#else\n#define V 12\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["12"]
    },
    if_nonzero_literal_is_true => {
        includes: ["<stdio.h>"],
        decls: "#if 42\n#define V 13\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["13"]
    },
    if_defined_and_expression => {
        includes: ["<stdio.h>"],
        decls: "#define X 1\n#define Y 1\n#if defined(X) && defined(Y)\n#define V 14\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["14"]
    },
    if_defined_or_expression => {
        includes: ["<stdio.h>"],
        decls: "#if defined(LEFT) || defined(RIGHT)\n#define V 0\n#else\n#define V 15\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["15"]
    },
    if_bitwise_and_mask => {
        includes: ["<stdio.h>"],
        decls: "#define FLAGS 5\n#if (FLAGS & 4)\n#define V 16\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["16"]
    },
    if_bitwise_or_mask => {
        includes: ["<stdio.h>"],
        decls: "#define FLAGS 1\n#if (FLAGS | 2)==3\n#define V 17\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["17"]
    },
    if_shift_in_expression => {
        includes: ["<stdio.h>"],
        decls: "#if (1<<3)==8\n#define V 18\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["18"]
    },
    if_modulo_in_expression => {
        includes: ["<stdio.h>"],
        decls: "#if (17%5)==2\n#define V 19\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["19"]
    },
    if_parenthesized_expression => {
        includes: ["<stdio.h>"],
        decls: "#define A 2\n#define B 3\n#if (A*B)+1==7\n#define V 21\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["21"]
    },
    elif_with_defined_check => {
        includes: ["<stdio.h>"],
        decls: "#define MODE 0\n#if MODE==1\n#define V 1\n#elif defined(MODE)\n#define V 22\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["22"]
    },
    if_macro_value_equality => {
        includes: ["<stdio.h>"],
        decls: "#define VERSION 3\n#if VERSION==3\n#define V 23\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["23"]
    },
    if_macro_value_inequality => {
        includes: ["<stdio.h>"],
        decls: "#define VERSION 3\n#if VERSION!=2\n#define V 24\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["24"]
    },
    conditional_pick_high_low => {
        includes: ["<stdio.h>"],
        decls: "#define HIGH_RES 1\n#if HIGH_RES\n#define DPI 300\n#else\n#define DPI 72\n#endif",
        body: "printf(\"%d\\n\", DPI); return 0;",
        expect: ["300"]
    },
    conditional_pick_low_res => {
        includes: ["<stdio.h>"],
        decls: "#undef HIGH_RES\n#if defined(HIGH_RES)\n#define DPI 300\n#else\n#define DPI 72\n#endif",
        body: "printf(\"%d\\n\", DPI); return 0;",
        expect: ["72"]
    },
    if_unary_not_on_zero => {
        includes: ["<stdio.h>"],
        decls: "#if !0\n#define V 25\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["25"]
    },
    if_unary_not_on_defined_missing => {
        includes: ["<stdio.h>"],
        decls: "#if !defined(ABSENT)\n#define V 26\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["26"]
    },
    if_chained_elif_numeric_ladder => {
        includes: ["<stdio.h>"],
        decls: "#define LEVEL 4\n#if LEVEL==1\n#define V 1\n#elif LEVEL==2\n#define V 2\n#elif LEVEL==3\n#define V 3\n#elif LEVEL==4\n#define V 27\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["27"]
    },
    if_expression_with_subtraction => {
        includes: ["<stdio.h>"],
        decls: "#define END 10\n#define START 3\n#if END-START==7\n#define V 28\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["28"]
    },
    if_expression_with_division => {
        includes: ["<stdio.h>"],
        decls: "#if 20/4==5\n#define V 29\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["29"]
    },
    ifdef_else_branch_value => {
        includes: ["<stdio.h>"],
        decls: "#ifdef PRESENT\n#define V 1\n#else\n#define V 30\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["30"]
    },
    ifndef_else_branch_value => {
        includes: ["<stdio.h>"],
        decls: "#define PRESENT\n#ifndef PRESENT\n#define V 0\n#else\n#define V 31\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["31"]
    },
    if_selects_between_two_constants => {
        includes: ["<stdio.h>"],
        decls: "#define USE_ALT 0\n#if USE_ALT\n#define PORT 8080\n#else\n#define PORT 80\n#endif",
        body: "printf(\"%d\\n\", PORT); return 0;",
        expect: ["80"]
    },
    if_selects_alternate_constant => {
        includes: ["<stdio.h>"],
        decls: "#define USE_ALT 1\n#if USE_ALT\n#define PORT 8080\n#else\n#define PORT 80\n#endif",
        body: "printf(\"%d\\n\", PORT); return 0;",
        expect: ["8080"]
    },
    nested_conditional_sum => {
        includes: ["<stdio.h>"],
        decls: "#define A 1\n#if A\n  #define X 10\n#else\n  #define X 0\n#endif\n#define B 1\n#if B\n  #define Y 22\n#else\n  #define Y 0\n#endif",
        body: "printf(\"%d\\n\", X+Y); return 0;",
        expect: ["32"]
    },
    if_comparison_on_negative_constant => {
        includes: ["<stdio.h>"],
        decls: "#define N -5\n#if N<0\n#define V 33\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["33"]
    },
    if_logical_and_short_circuit_value => {
        includes: ["<stdio.h>"],
        decls: "#define A 1\n#define B 0\n#if A && B\n#define V 1\n#else\n#define V 34\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["34"]
    },
    if_logical_or_enables_branch => {
        includes: ["<stdio.h>"],
        decls: "#define A 0\n#define B 1\n#if A || B\n#define V 35\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["35"]
    },
    if_defined_self_reference_guard => {
        includes: ["<stdio.h>"],
        decls: "#ifndef INC_ONCE\n#define INC_ONCE\n#define V 36\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["36"]
    },
    if_tertiary_via_macros => {
        includes: ["<stdio.h>"],
        decls: "#define DEBUG 0\n#if DEBUG\n#define LOG_LEVEL 3\n#elif 1\n#define LOG_LEVEL 1\n#else\n#define LOG_LEVEL 0\n#endif",
        body: "printf(\"%d\\n\", LOG_LEVEL); return 0;",
        expect: ["1"]
    },
    if_hex_constant_comparison => {
        includes: ["<stdio.h>"],
        decls: "#if 0x10==16\n#define V 37\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["37"]
    },
    if_octal_constant_comparison => {
        includes: ["<stdio.h>"],
        decls: "#if 010==8\n#define V 38\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["38"]
    },
    if_multiplication_precedence => {
        includes: ["<stdio.h>"],
        decls: "#if 2+3*4==14\n#define V 39\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["39"]
    },
    if_defined_complement_on_set_macro => {
        includes: ["<stdio.h>"],
        decls: "#define READY\n#if defined(READY)\n#define V 40\n#else\n#define V 0\n#endif",
        body: "printf(\"%d\\n\", V); return 0;",
        expect: ["40"]
    },
}

c_compile_cases! {
    if_false_branch_code_omitted_compile => {
        includes: ["<stdio.h>"],
        decls: "#if 0\nvoid dead(void);\n#endif\nint x=1;",
        body: "return x;"
    },
    elif_nested_compile => {
        includes: ["<stdio.h>"],
        decls: "#if 0\nint a=1;\n#elif 1\nint a=2;\n#endif",
        body: "return a;"
    },
}
