//! Preprocessor macro expansion — object/function macros, stringify, paste, variadic.

c_run_cases! {
    object_macro_integer_substitution_in_printf => {
        includes: ["<stdio.h>"],
        decls: "#define ANSWER 42",
        body: "printf(\"%d\\n\", ANSWER); return 0;",
        expect: ["42"]
    },
    object_macro_expression_operand => {
        includes: ["<stdio.h>"],
        decls: "#define BASE 10",
        body: "printf(\"%d\\n\", BASE * 3); return 0;",
        expect: ["30"]
    },
    function_macro_add_two_integers => {
        includes: ["<stdio.h>"],
        decls: "#define ADD(a,b) ((a)+(b))",
        body: "printf(\"%d\\n\", ADD(11, 22)); return 0;",
        expect: ["33"]
    },
    function_macro_preserves_precedence_with_parens => {
        includes: ["<stdio.h>"],
        decls: "#define DOUBLE(x) ((x)*2)",
        body: "printf(\"%d\\n\", DOUBLE(3+4)); return 0;",
        expect: ["14"]
    },
    function_macro_square_with_parens => {
        includes: ["<stdio.h>"],
        decls: "#define SQUARE(x) ((x)*(x))",
        body: "printf(\"%d\\n\", SQUARE(1+2)); return 0;",
        expect: ["9"]
    },
    stringify_operator_in_printf => {
        includes: ["<stdio.h>"],
        decls: "#define STR(x) #x",
        body: "printf(\"%s\\n\", STR(hello)); return 0;",
        expect: ["hello"]
    },
    stringify_operator_with_integer_token => {
        includes: ["<stdio.h>"],
        decls: "#define STR(x) #x",
        body: "printf(\"%s\\n\", STR(123)); return 0;",
        expect: ["123"]
    },
    token_paste_creates_identifier => {
        includes: ["<stdio.h>"],
        decls: "#define CAT(a,b) a##b\nint xy = 55;",
        body: "printf(\"%d\\n\", CAT(x,y)); return 0;",
        expect: ["55"]
    },
    token_paste_prefix_suffix_variable => {
        includes: ["<stdio.h>"],
        decls: "#define PREFIX(n) val_##n\nint val_7 = 70;",
        body: "printf(\"%d\\n\", PREFIX(7)); return 0;",
        expect: ["70"]
    },
    variadic_macro_forwards_printf_args => {
        includes: ["<stdio.h>"],
        decls: "#define LOG(fmt, ...) printf(fmt, __VA_ARGS__)",
        body: "LOG(\"%d %s\\n\", 9, \"ok\"); return 0;",
        expect: ["9 ok"]
    },
    variadic_macro_single_extra_arg => {
        includes: ["<stdio.h>"],
        decls: "#define P1(fmt, a) printf(fmt, a)",
        body: "P1(\"%d\\n\", 17); return 0;",
        expect: ["17"]
    },
    nested_macro_expansion_order => {
        includes: ["<stdio.h>"],
        decls: "#define TWO 2\n#define FOUR (TWO*TWO)",
        body: "printf(\"%d\\n\", FOUR); return 0;",
        expect: ["4"]
    },
    macro_chain_triple_expand => {
        includes: ["<stdio.h>"],
        decls: "#define A 1\n#define B (A+A)\n#define C (B+B)",
        body: "printf(\"%d\\n\", C); return 0;",
        expect: ["4"]
    },
    macro_in_macro_argument_expands_first => {
        includes: ["<stdio.h>"],
        decls: "#define X 5\n#define TWICE(n) ((n)*2)",
        body: "printf(\"%d\\n\", TWICE(X)); return 0;",
        expect: ["10"]
    },
    multiline_macro_swap_values => {
        includes: ["<stdio.h>"],
        decls: "#define SWAP(a,b) do{int _t=a;a=b;b=_t;}while(0)",
        body: "int x=1,y=9; SWAP(x,y); printf(\"%d %d\\n\", x, y); return 0;",
        expect: ["9 1"]
    },
    undef_then_redefine_changes_value => {
        includes: ["<stdio.h>"],
        decls: "#define N 3\n#undef N\n#define N 8",
        body: "printf(\"%d\\n\", N); return 0;",
        expect: ["8"]
    },
    macro_expands_to_string_literal => {
        includes: ["<stdio.h>"],
        decls: "#define GREET \"hi\"",
        body: "printf(\"%s\\n\", GREET); return 0;",
        expect: ["hi"]
    },
    macro_expands_to_char_literal => {
        includes: ["<stdio.h>"],
        decls: "#define CH 'Q'",
        body: "printf(\"%c\\n\", CH); return 0;",
        expect: ["Q"]
    },
    macro_used_in_array_dimension => {
        includes: ["<stdio.h>"],
        decls: "#define LEN 3",
        body: "int a[LEN]={1,2,3}; printf(\"%d\\n\", a[LEN-1]); return 0;",
        expect: ["3"]
    },
    macro_expands_inside_switch_case => {
        includes: ["<stdio.h>"],
        decls: "#define TAG 2",
        body: "switch(2){case TAG: printf(\"hit\\n\"); break; default: printf(\"miss\\n\");} return 0;",
        expect: ["hit"]
    },
    macro_comparison_in_expression => {
        includes: ["<stdio.h>"],
        decls: "#define GT(x,y) ((x)>(y))",
        body: "printf(\"%d\\n\", GT(8,3)); return 0;",
        expect: ["1"]
    },
    macro_bit_shift_mask => {
        includes: ["<stdio.h>"],
        decls: "#define BIT(n) (1<<(n))",
        body: "printf(\"%d\\n\", BIT(4)); return 0;",
        expect: ["16"]
    },
    macro_stringify_after_concat => {
        includes: ["<stdio.h>"],
        decls: "#define MK(n) val_##n\n#define STR(x) #x\nint val_x = 1;",
        body: "printf(\"%s\\n\", STR(MK(x))); return 0;",
        expect: ["val_x"]
    },
    macro_paste_with_numeric_suffix => {
        includes: ["<stdio.h>"],
        decls: "#define IDX(n) item##n\nint item3 = 33;",
        body: "printf(\"%d\\n\", IDX(3)); return 0;",
        expect: ["33"]
    },
    variadic_macro_empty_va_args => {
        includes: ["<stdio.h>"],
        decls: "#define SHOW(fmt, ...) printf(fmt, ##__VA_ARGS__)",
        body: "SHOW(\"done\\n\"); return 0;",
        expect: ["done"]
    },
    macro_expands_to_float_literal => {
        includes: ["<stdio.h>"],
        decls: "#define HALF 0.5",
        body: "printf(\"%.1f\\n\", HALF + HALF); return 0;",
        expect: ["1.0"]
    },
    macro_expands_to_cast_expression => {
        includes: ["<stdio.h>"],
        decls: "#define TO_INT(x) ((int)(x))",
        body: "printf(\"%d\\n\", TO_INT(3.9)); return 0;",
        expect: ["3"]
    },
    macro_for_loop_upper_bound => {
        includes: ["<stdio.h>"],
        decls: "#define LAST 2",
        body: "for(int i=0;i<=LAST;i++) if(i==LAST) printf(\"%d\\n\", i); return 0;",
        expect: ["2"]
    },
    macro_logical_and_combination => {
        includes: ["<stdio.h>"],
        decls: "#define BOTH(a,b) ((a)&&(b))",
        body: "printf(\"%d\\n\", BOTH(1,0)); return 0;",
        expect: ["0"]
    },
    macro_ternary_selection => {
        includes: ["<stdio.h>"],
        decls: "#define PICK(c,a,b) ((c)?(a):(b))",
        body: "printf(\"%d\\n\", PICK(0, 7, 4)); return 0;",
        expect: ["4"]
    },
    macro_max_of_two => {
        includes: ["<stdio.h>"],
        decls: "#define MAX(a,b) ((a)>(b)?(a):(b))",
        body: "printf(\"%d\\n\", MAX(12, 8)); return 0;",
        expect: ["12"]
    },
    macro_min_of_two => {
        includes: ["<stdio.h>"],
        decls: "#define MIN(a,b) ((a)<(b)?(a):(b))",
        body: "printf(\"%d\\n\", MIN(12, 8)); return 0;",
        expect: ["8"]
    },
    macro_abs_value => {
        includes: ["<stdio.h>"],
        decls: "#define ABS(x) ((x)<0?-(x):(x))",
        body: "printf(\"%d\\n\", ABS(-15)); return 0;",
        expect: ["15"]
    },
    macro_concat_function_name => {
        includes: ["<stdio.h>"],
        decls: "#define CAT(a,b) a##b\nint fn42(void){return 42;}",
        body: "printf(\"%d\\n\", CAT(fn,42)()); return 0;",
        expect: ["42"]
    },
    stringify_operator_with_macro_arg => {
        includes: ["<stdio.h>"],
        decls: "#define X foo\n#define STR(x) #x",
        body: "printf(\"%s\\n\", STR(X)); return 0;",
        expect: ["foo"]
    },
    macro_recursive_style_double => {
        includes: ["<stdio.h>"],
        decls: "#define INC(x) ((x)+1)\n#define TWICE(x) INC(INC(x))",
        body: "printf(\"%d\\n\", TWICE(3)); return 0;",
        expect: ["5"]
    },
    macro_expands_in_printf_format_width => {
        includes: ["<stdio.h>"],
        decls: "#define W 5",
        body: "printf(\"%*d\\n\", W, 7); return 0;",
        expect: ["    7"]
    },
    macro_comma_operator_sequence => {
        includes: ["<stdio.h>"],
        decls: "#define SEQ(a,b) ((a),(b))",
        body: "printf(\"%d\\n\", SEQ(1,2)); return 0;",
        expect: ["2"]
    },
    macro_expands_multiple_printf_lines => {
        includes: ["<stdio.h>"],
        decls: "#define PAIR printf(\"%d\\n\",1); printf(\"%d\\n\",2)",
        body: "PAIR; return 0;",
        expect: ["1", "2"]
    },
    macro_stringize_spaces_preserved => {
        includes: ["<stdio.h>"],
        decls: "#define STR(x) #x",
        body: "printf(\"%s\\n\", STR(a b)); return 0;",
        expect: ["a b"]
    },
    macro_function_three_args => {
        includes: ["<stdio.h>"],
        decls: "#define SUM3(a,b,c) ((a)+(b)+(c))",
        body: "printf(\"%d\\n\", SUM3(1,2,3)); return 0;",
        expect: ["6"]
    },
    macro_expands_in_initializer => {
        includes: ["<stdio.h>"],
        decls: "#define SEED 4",
        body: "int v=SEED+1; printf(\"%d\\n\", v); return 0;",
        expect: ["5"]
    },
    macro_paste_in_printf_identifier => {
        includes: ["<stdio.h>"],
        decls: "#define VAR(n) x##n\nint x9=99;",
        body: "printf(\"%d\\n\", VAR(9)); return 0;",
        expect: ["99"]
    },
    variadic_macro_three_values => {
        includes: ["<stdio.h>"],
        decls: "#define EMIT(fmt, ...) printf(fmt, __VA_ARGS__)",
        body: "EMIT(\"%d %d %d\\n\", 1, 2, 3); return 0;",
        expect: ["1 2 3"]
    },
    macro_negation_of_constant => {
        includes: ["<stdio.h>"],
        decls: "#define VAL 8\n#define NEG(x) (-(x))",
        body: "printf(\"%d\\n\", NEG(VAL)); return 0;",
        expect: ["-8"]
    },
    macro_modulo_expression => {
        includes: ["<stdio.h>"],
        decls: "#define MOD(a,b) ((a)%(b))",
        body: "printf(\"%d\\n\", MOD(17, 5)); return 0;",
        expect: ["2"]
    },
    macro_expands_in_return_path_guard => {
        includes: ["<stdio.h>"],
        decls: "#define OK 1",
        body: "if(OK) printf(\"yes\\n\"); else printf(\"no\\n\"); return 0;",
        expect: ["yes"]
    },
}

c_compile_cases! {
    macro_multiline_backslash_compile => {
        includes: ["<stdio.h>"],
        decls: "#define INC(x) \\\n((x)+1)",
        body: "return INC(1);"
    },
    macro_paste_in_declaration_compile => {
        includes: ["<stdio.h>"],
        decls: "#define TYPEDEF_NAME counter\nint TYPEDEF_NAME = 0;",
        body: "return TYPEDEF_NAME;"
    },
}
