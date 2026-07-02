//! snprintf, vsnprintf, and formatted buffer I/O — distinct format/width cases.


c_run_cases! {
    snprintf_writes_null_terminated => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%s\",\"hi\"); printf(\"%s\\n\", b); return 0;",
        expect: ["hi"]
    },
    snprintf_truncates_with_n => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[4]; snprintf(b,4,\"%s\",\"abcde\"); printf(\"%s\\n\", b); return 0;",
        expect: ["abc"]
    },
    snprintf_returns_char_count => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[4]; int n=snprintf(b,4,\"%s\",\"abcde\"); printf(\"%d\\n\", n); return 0;",
        expect: ["5"]
    },
    snprintf_zero_size_no_write => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[4]={'z','z','z','z'}; snprintf(b,0,\"%s\",\"x\"); printf(\"%c\\n\", b[0]); return 0;",
        expect: ["z"]
    },
    snprintf_integer_padding => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%04d\",7); printf(\"%s\\n\", b); return 0;",
        expect: ["0007"]
    },
    snprintf_negative_width => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%-4d\",9); printf(\"%s\\n\", b); return 0;",
        expect: ["9   "]
    },
    snprintf_hex_upper => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%02X\",15); printf(\"%s\\n\", b); return 0;",
        expect: ["0F"]
    },
    snprintf_octal_output => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%o\",8); printf(\"%s\\n\", b); return 0;",
        expect: ["10"]
    },
    snprintf_unsigned_decimal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[12]; snprintf(b,12,\"%u\",4000000000u); printf(\"%s\\n\", b); return 0;",
        expect: ["4000000000"]
    },
    snprintf_long_long => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[16]; snprintf(b,16,\"%lld\",-123456789012LL); printf(\"%s\\n\", b); return 0;",
        expect: ["-123456789012"]
    },
    snprintf_size_t_zu => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%zu\",(size_t)12); printf(\"%s\\n\", b); return 0;",
        expect: ["12"]
    },
    snprintf_pointer_percent_p => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[32]; int x; snprintf(b,32,\"%p\",(void*)&x); printf(\"%d\\n\", b[0]=='0'); return 0;",
        expect: ["1"]
    },
    snprintf_double_precision => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[16]; snprintf(b,16,\"%.3f\",1.2345); printf(\"%s\\n\", b); return 0;",
        expect: ["1.235"]
    },
    snprintf_scientific_notation => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[16]; snprintf(b,16,\"%.1e\",250.0); printf(\"%s\\n\", b); return 0;",
        expect: ["2.5e+02"]
    },
    snprintf_multiple_fields => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[16]; snprintf(b,16,\"%d-%s\",3,\"ok\"); printf(\"%s\\n\", b); return 0;",
        expect: ["3-ok"]
    },
    snprintf_percent_literal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"50%%\"); printf(\"%s\\n\", b); return 0;",
        expect: ["50%"]
    },
    snprintf_char_conversion => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%c\",66); printf(\"%s\\n\", b); return 0;",
        expect: ["B"]
    },
    snprintf_string_precision => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%.2s\",\"four\"); printf(\"%s\\n\", b); return 0;",
        expect: ["fo"]
    },
    snprintf_empty_format => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[4]={'x','x','x','\\0'}; snprintf(b,4,\"\"); printf(\"%d\\n\", b[0]=='x'); return 0;",
        expect: ["1"]
    },
    snprintf_plus_sign_flag => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%+d\",5); printf(\"%s\\n\", b); return 0;",
        expect: ["+5"]
    },
    snprintf_space_sign_flag => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"% d\",5); printf(\"%s\\n\", b); return 0;",
        expect: [" 5"]
    },
    snprintf_alternate_hex => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%#x\",10); printf(\"%s\\n\", b); return 0;",
        expect: ["0xa"]
    },
    snprintf_short_int_hd => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%hd\",(short)-3); printf(\"%s\\n\", b); return 0;",
        expect: ["-3"]
    },
    snprintf_long_hex_lx => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[12]; snprintf(b,12,\"%lx\",255L); printf(\"%s\\n\", b); return 0;",
        expect: ["ff"]
    },
    vsnprintf_basic => {
        includes: ["<stdio.h>", "<stdarg.h>"],
        decls: "int fmt(char *b,int n,const char *f,...){va_list ap; va_start(ap,f); int r=vsnprintf(b,n,f,ap); va_end(ap); return r;}",
        body: "char b[8]; fmt(b,8,\"%d\",42); printf(\"%s\\n\", b); return 0;",
        expect: ["42"]
    },
    vsnprintf_truncation_return => {
        includes: ["<stdio.h>", "<stdarg.h>"],
        decls: "int fmt(char *b,int n,const char *f,...){va_list ap; va_start(ap,f); int r=vsnprintf(b,n,f,ap); va_end(ap); return r;}",
        body: "char b[4]; int n=fmt(b,4,\"%s\",\"long\"); printf(\"%d\\n\", n); return 0;",
        expect: ["4"]
    },
    sprintf_writes_to_buffer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; sprintf(b,\"%d\",11); printf(\"%s\\n\", b); return 0;",
        expect: ["11"]
    },
    sprintf_multiple_args => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[16]; sprintf(b,\"%c%d\",\"Z\",9); printf(\"%s\\n\", b); return 0;",
        expect: ["Z9"]
    },
    fputs_to_stdout => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "fputs(\"line\\n\", stdout); return 0;",
        expect: ["line"]
    },
    fputs_without_newline => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "fputs(\"xy\", stdout); fputs(\"\\n\", stdout); return 0;",
        expect: ["xy"]
    },
    fputc_writes_char => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "fputc('Q', stdout); fputc('\\n', stdout); return 0;",
        expect: ["Q"]
    },
    putchar_writes => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "putchar('M'); putchar('\\n'); return 0;",
        expect: ["M"]
    },
    puts_appends_newline => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "puts(\"z\"); return 0;",
        expect: ["z"]
    },
    fputs_return_nonnegative => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", fputs(\"a\", stdout) >= 0); return 0;",
        expect: ["1"]
    },
    snprintf_into_stack_array => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[6]; snprintf(b,6,\"%d+%d\",2,3); printf(\"%s\\n\", b); return 0;",
        expect: ["2+3"]
    },
    snprintf_width_string => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[10]; snprintf(b,10,\"|%5s|\",\"go\"); printf(\"%s\\n\", b); return 0;",
        expect: ["|   go|"]
    },
    snprintf_zero_value => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%d\",0); printf(\"%s\\n\", b); return 0;",
        expect: ["0"]
    },
    snprintf_negative_float => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[12]; snprintf(b,12,\"%.1f\",-2.5); printf(\"%s\\n\", b); return 0;",
        expect: ["-2.5"]
    },
    snprintf_general_format => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[12]; snprintf(b,12,\"%.2g\",12.3); printf(\"%s\\n\", b); return 0;",
        expect: ["12"]
    },
    snprintf_i_conversion => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%i\",-8); printf(\"%s\\n\", b); return 0;",
        expect: ["-8"]
    },
    snprintf_u_with_width => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%3u\",7u); printf(\"%s\\n\", b); return 0;",
        expect: ["  7"]
    },
    snprintf_x_lowercase => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%x\",255); printf(\"%s\\n\", b); return 0;",
        expect: ["ff"]
    },
    snprintf_x_uppercase => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[8]; snprintf(b,8,\"%X\",255); printf(\"%s\\n\", b); return 0;",
        expect: ["FF"]
    },
    snprintf_null_string_guard => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[12]; snprintf(b,12,\"%s\",(char*)0); printf(\"%d\\n\", b[0]=='('); return 0;",
        expect: ["1"]
    },
    snprintf_repeat_small_buffer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[3]; snprintf(b,3,\"%d%d\",1,2); printf(\"%s\\n\", b); return 0;",
        expect: ["12"]
    },
    snprintf_ptrdiff_td => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "",
        body: "char b[12]; snprintf(b,12,\"%td\",(ptrdiff_t)-4); printf(\"%s\\n\", b); return 0;",
        expect: ["-4"]
    },
    snprintf_uintmax_ju => {
        includes: ["<stdio.h>", "<stdint.h>", "<inttypes.h>"],
        decls: "",
        body: "char b[16]; snprintf(b,16,\"%ju\",(uintmax_t)100); printf(\"%s\\n\", b); return 0;",
        expect: ["100"]
    },
    snprintf_intmax_jd => {
        includes: ["<stdio.h>", "<stdint.h>", "<inttypes.h>"],
        decls: "",
        body: "char b[16]; snprintf(b,16,\"%jd\",(intmax_t)-55); printf(\"%s\\n\", b); return 0;",
        expect: ["-55"]
    },
    snprintf_long_double_lf => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char b[16]; snprintf(b,16,\"%.1Lf\",3.5L); printf(\"%s\\n\", b); return 0;",
        expect: ["3.5"]
    },
}
