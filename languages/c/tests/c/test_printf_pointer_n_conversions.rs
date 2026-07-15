//! printf pointer, %n, and %% edge cases beyond test_printf_formats.rs.

c_run_cases! {
    printf_n_records_three_chars => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"abc%n\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["abc3"]
    },
    printf_n_after_integer => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"x%d%n\", 7, &n); printf(\"%d\\n\", n); return 0;",
        expect: ["x72"]
    },
    printf_n_zero_width_output => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=99; printf(\"%n\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["0"]
    },
    printf_n_after_percent_literal => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"50%%%n\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["50%3"]
    },
    printf_n_mid_format_string => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int a=0,b=0; printf(\"hi%n there%n\", &a,&b); printf(\"%d %d\\n\", a,b); return 0;",
        expect: ["hi there2 8"]
    },
    printf_n_with_string_conversion => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%s%n\", \"go\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["go2"]
    },
    printf_n_after_float => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%.1f%n\", 1.5, &n); printf(\"%d\\n\", n); return 0;",
        expect: ["1.53"]
    },
    printf_n_multiple_separate_prints => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"a%n\", &n); int m=n; printf(\"bc%n\", &n); printf(\"%d %d\\n\", m, n); return 0;",
        expect: ["abc1 3"]
    },
    printf_n_before_newline => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%d%n\\n\", 5, &n); printf(\"%d\\n\", n); return 0;",
        expect: ["5", "2"]
    },
    printf_n_after_width_field => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%5d%n\", 1, &n); printf(\"%d\\n\", n); return 0;",
        expect: ["    15"]
    },
    printf_p_null_void_pointer => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%p\\n\", (void*)0); return 0;",
        expect: ["0x0"]
    },
    printf_p_null_int_pointer => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int *p=0; printf(\"%p\\n\", (void*)p); return 0;",
        expect: ["0x0"]
    },
    printf_p_same_address_printed_twice => {
        includes: ["<stdio.h>", "<stddef.h>", "<string.h>"],
        decls: "",
        body: "char a[64],b[64]; int x; void *p=&x; sprintf(a,\"%p\",p); sprintf(b,\"%p\",p); printf(\"%d\\n\", strcmp(a,b)==0); return 0;",
        expect: ["1"]
    },
    printf_p_void_pointer_from_int_address => {
        includes: ["<stdio.h>", "<stddef.h>", "<string.h>"],
        decls: "",
        body: "char a[64],b[64]; int x=4; void *p=&x; sprintf(a,\"%p\",p); sprintf(b,\"%p\",p); printf(\"%d\\n\", strcmp(a,b)==0); return 0;",
        expect: ["1"]
    },
    printf_p_array_decayed_pointer => {
        includes: ["<stdio.h>", "<stddef.h>", "<string.h>"],
        decls: "",
        body: "char a[64],b[64]; int arr[2]={3,4}; sprintf(a,\"%p\",(void*)arr); sprintf(b,\"%p\",(void*)arr); printf(\"%d\\n\", strcmp(a,b)==0); return 0;",
        expect: ["1"]
    },
    printf_p_char_array_address => {
        includes: ["<stdio.h>", "<stddef.h>", "<string.h>"],
        decls: "",
        body: "char a[64],b[64]; char s[]=\"Z\"; sprintf(a,\"%p\",(void*)s); sprintf(b,\"%p\",(void*)s); printf(\"%d\\n\", strcmp(a,b)==0); return 0;",
        expect: ["1"]
    },
    printf_p_with_literal_prefix => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"ptr=%p\\n\", (void*)0); return 0;",
        expect: ["ptr=0x0"]
    },
    printf_p_width_field_on_null => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%12p\\n\", (void*)0); return 0;",
        expect: ["         0x0"]
    },
    printf_p_const_data_pointer => {
        includes: ["<stdio.h>", "<stddef.h>", "<string.h>"],
        decls: "",
        body: "char a[64],b[64]; const int x=9; const void *p=&x; sprintf(a,\"%p\",p); sprintf(b,\"%p\",p); printf(\"%d\\n\", strcmp(a,b)==0); return 0;",
        expect: ["1"]
    },
    printf_p_static_variable_address => {
        includes: ["<stdio.h>", "<stddef.h>", "<string.h>"],
        decls: "",
        body: "char a[64],b[64]; static int g=11; sprintf(a,\"%p\",(void*)&g); sprintf(b,\"%p\",(void*)&g); printf(\"%d\\n\", strcmp(a,b)==0); return 0;",
        expect: ["1"]
    },
    printf_p_two_distinct_stack_objects_differ => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int a,b; printf(\"%d\\n\", (void*)&a != (void*)&b); return 0;",
        expect: ["1"]
    },
    printf_p_function_pointer_type => {
        includes: ["<stdio.h>", "<stddef.h>", "<string.h>"],
        decls: "int id(int v){ return v; }",
        body: "char a[64],b[64]; sprintf(a,\"%p\",(void*)id); sprintf(b,\"%p\",(void*)id); printf(\"%d\\n\", strcmp(a,b)==0); return 0;",
        expect: ["1"]
    },
    printf_percent_doubled_at_start => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%%start\\n\"); return 0;",
        expect: ["%start"]
    },
    printf_percent_doubled_at_end => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"end%%\\n\"); return 0;",
        expect: ["end%"]
    },
    printf_percent_triple_emits_two => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"a%%%b\\n\"); return 0;",
        expect: ["a%b"]
    },
    printf_percent_quad_emits_two_literals => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%%%%\\n\"); return 0;",
        expect: ["%%"]
    },
    printf_percent_between_ints => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%d%% %d\\n\", 1, 2); return 0;",
        expect: ["1% 2"]
    },
    printf_percent_before_string => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%%s is literal\\n\"); return 0;",
        expect: ["%s is literal"]
    },
    printf_percent_after_string => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"ratio %s%%\\n\", \"50\"); return 0;",
        expect: ["ratio 50%"]
    },
    printf_percent_with_width_on_literal => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%5%%\\n\"); return 0;",
        expect: ["    %"]
    },
    printf_percent_mixed_with_hex => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"0x%x%%\\n\", 10); return 0;",
        expect: ["0xa%"]
    },
    printf_percent_only_output => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%%\\n\"); return 0;",
        expect: ["%"]
    },
    printf_percent_in_loop => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "for(int i=0;i<3;i++) printf(\"%%\"); printf(\"\\n\"); return 0;",
        expect: ["%%%"]
    },
    printf_percent_with_float => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%.0f%%\\n\", 99.5); return 0;",
        expect: ["100%"]
    },
    printf_percent_adjacent_to_char => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%c%%\\n\", 65); return 0;",
        expect: ["A%"]
    },
    printf_percent_after_octal => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%o%%\\n\", 8); return 0;",
        expect: ["10%"]
    },
    printf_percent_after_unsigned => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%u%%\\n\", 3u); return 0;",
        expect: ["3%"]
    },
    printf_percent_left_justified => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%-4%%|\\n\"); return 0;",
        expect: ["%   |"]
    },
    printf_percent_zero_padded_width => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%03%%\\n\"); return 0;",
        expect: ["00%"]
    },
    printf_percent_between_spaces => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\" %%% \\n\"); return 0;",
        expect: [" % "]
    },
    printf_percent_with_newline_escape => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"done%%\\n\"); return 0;",
        expect: ["done%"]
    },
    printf_percent_and_n_combo => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%% %n\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["% 2"]
    },
    printf_n_after_wide_string_field => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%10s%n\", \"x\", &n); printf(\"%d\\n\", n); return 0;",
        expect: ["         x10"]
    },
    printf_n_with_zero_precision_int => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%.0d%n\", 0, &n); printf(\"%d\\n\", n); return 0;",
        expect: ["0"]
    },
    printf_percent_star_width => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%*s%%\\n\", 3, \"a\"); return 0;",
        expect: ["  a%"]
    },
    printf_percent_in_format_tail => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"100 %%% done\\n\"); return 0;",
        expect: ["100 % done"]
    },
    printf_p_pointer_chain => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int x=2; int *p=&x; int **pp=&p; printf(\"%d\\n\", **pp); return 0;",
        expect: ["2"]
    },
    printf_n_after_multiple_flags => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%+05d%n\", 3, &n); printf(\"%d\\n\", n); return 0;",
        expect: ["+00035"]
    },
    printf_percent_hex_upper_with_suffix => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%X%%\\n\", 255); return 0;",
        expect: ["FF%"]
    },
    printf_percent_scientific_suffix => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%.0e%%\\n\", 1000.0); return 0;",
        expect: ["1e+03%"]
    },
    printf_p_struct_member_via_pointer => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "struct S{int a,b;}; struct S s={1,2}; int *p=&s.b; printf(\"%d\\n\", *p); return 0;",
        expect: ["2"]
    },
    printf_n_after_char_conversion => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "int n=0; printf(\"%c%n\", 65, &n); printf(\"%d\\n\", n); return 0;",
        expect: ["A1"]
    },
    printf_percent_general_format_suffix => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "printf(\"%.1g%%\\n\", 12.3); return 0;",
        expect: ["1e+01%"]
    },
}
