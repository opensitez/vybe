//! printf length modifiers — %hd %hhd %ld %lld %zu %td %jd %Lf and related.


c_run_cases! {
    printf_hhd_unsigned_char_value => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned char uc=200; printf(\"%hhu\\n\", uc); return 0;",
        expect: ["200"]
    },
    printf_hhd_signed_char_negative => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "signed char sc=-5; printf(\"%hhd\\n\", sc); return 0;",
        expect: ["-5"]
    },
    printf_hhd_char_zero => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "signed char sc=0; printf(\"%hhd\\n\", sc); return 0;",
        expect: ["0"]
    },
    printf_hhd_char_positive => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "signed char sc=42; printf(\"%hhd\\n\", sc); return 0;",
        expect: ["42"]
    },
    printf_hd_short_max => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "short s=32767; printf(\"%hd\\n\", s); return 0;",
        expect: ["32767"]
    },
    printf_hd_short_negative => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "short s=-1234; printf(\"%hd\\n\", s); return 0;",
        expect: ["-1234"]
    },
    printf_hu_unsigned_short => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned short us=65000; printf(\"%hu\\n\", us); return 0;",
        expect: ["65000"]
    },
    printf_hu_small_unsigned_short => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned short us=255; printf(\"%hu\\n\", us); return 0;",
        expect: ["255"]
    },
    printf_ld_negative_long => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "long v=-999999L; printf(\"%ld\\n\", v); return 0;",
        expect: ["-999999"]
    },
    printf_ld_positive_long => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "long v=500000L; printf(\"%ld\\n\", v); return 0;",
        expect: ["500000"]
    },
    printf_lu_large_unsigned_long => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long v=5000000000UL; printf(\"%lu\\n\", v); return 0;",
        expect: ["5000000000"]
    },
    printf_lu_small_unsigned_long => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long v=42UL; printf(\"%lu\\n\", v); return 0;",
        expect: ["42"]
    },
    printf_lld_negative_longlong => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "long long v=-9223372036854775807LL; printf(\"%lld\\n\", v); return 0;",
        expect: ["-9223372036854775807"]
    },
    printf_lld_positive_longlong => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "long long v=1234567890123LL; printf(\"%lld\\n\", v); return 0;",
        expect: ["1234567890123"]
    },
    printf_llu_unsigned_longlong => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long long v=18446744073709551615ULL; printf(\"%llu\\n\", v); return 0;",
        expect: ["18446744073709551615"]
    },
    printf_llu_moderate_value => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long long v=10000000000ULL; printf(\"%llu\\n\", v); return 0;",
        expect: ["10000000000"]
    },
    printf_zu_sizeof_char => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "printf(\"%zu\\n\", sizeof(char)); return 0;",
        expect: ["1"]
    },
    printf_zu_sizeof_double => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "printf(\"%zu\\n\", sizeof(double)); return 0;",
        expect: ["8"]
    },
    printf_zu_sizeof_pointer => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "printf(\"%zu\\n\", sizeof(void*)); return 0;",
        expect: ["8"]
    },
    printf_zu_array_total_bytes => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "int arr[5]; printf(\"%zu\\n\", sizeof arr); return 0;",
        expect: ["20"]
    },
    printf_td_positive_ptrdiff => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "int a[4]; ptrdiff_t d=&a[3]-&a[0]; printf(\"%td\\n\", d); return 0;",
        expect: ["3"]
    },
    printf_td_negative_ptrdiff => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "int a[4]; ptrdiff_t d=&a[0]-&a[2]; printf(\"%td\\n\", d); return 0;",
        expect: ["-2"]
    },
    printf_td_zero_ptrdiff => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "int x; ptrdiff_t d=&x-&x; printf(\"%td\\n\", d); return 0;",
        expect: ["0"]
    },
    printf_jd_intmax_positive => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "intmax_t v=9223372036854775807; printf(\"%jd\\n\", v); return 0;",
        expect: ["9223372036854775807"]
    },
    printf_jd_intmax_negative => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "intmax_t v=-42; printf(\"%jd\\n\", v); return 0;",
        expect: ["-42"]
    },
    printf_ju_uintmax_value => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "uintmax_t v=18446744073709551615U; printf(\"%ju\\n\", v); return 0;",
        expect: ["18446744073709551615"]
    },
    printf_lx_long_hex => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long v=0xdeadbeefUL; printf(\"%lx\\n\", v); return 0;",
        expect: ["deadbeef"]
    },
    printf_llx_longlong_hex => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long long v=0x1234ULL; printf(\"%llx\\n\", v); return 0;",
        expect: ["1234"]
    },
    printf_llo_long_octal => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long v=511UL; printf(\"%llo\\n\", v); return 0;",
        expect: ["777"]
    },
    printf_lf_double_default => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "double d=2.5; printf(\"%lf\\n\", d); return 0;",
        expect: ["2.500000"]
    },
    printf_lf_long_double_literal => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "long double ld=1.25L; printf(\"%Lf\\n\", ld); return 0;",
        expect: ["1.250000"]
    },
    printf_hd_width_and_sign => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "short s=7; printf(\"%+5hd\\n\", s); return 0;",
        expect: ["   +7"]
    },
    printf_hhd_zero_padded => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "signed char c=5; printf(\"%03hhd\\n\", c); return 0;",
        expect: ["005"]
    },
    printf_ld_width_eight => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "long v=42L; printf(\"%8ld\\n\", v); return 0;",
        expect: ["      42"]
    },
    printf_lld_alternate_hex => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long long v=255ULL; printf(\"%#llx\\n\", v); return 0;",
        expect: ["0xff"]
    },
    printf_zu_left_justified => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "size_t s=99; printf(\"%-5zu|\\n\", s); return 0;",
        expect: ["99   |"]
    },
    printf_td_with_plus_flag => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "ptrdiff_t d=4; printf(\"%+td\\n\", d); return 0;",
        expect: ["+4"]
    },
    printf_hu_octal => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned short us=8; printf(\"%ho\\n\", us); return 0;",
        expect: ["10"]
    },
    printf_hd_hex_upper => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "short s=255; printf(\"%hX\\n\", s); return 0;",
        expect: ["FF"]
    },
    printf_lld_negative_decimal => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "long long v=-500LL; printf(\"%lld\\n\", v); return 0;",
        expect: ["-500"]
    },
    printf_lu_width_twelve => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long v=7UL; printf(\"%12lu\\n\", v); return 0;",
        expect: ["           7"]
    },
    printf_ju_hex_lower => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "uintmax_t v=255; printf(\"%jx\\n\", v); return 0;",
        expect: ["ff"]
    },
    printf_zu_precision_field => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "size_t s=0; printf(\"%.0zu\\n\", s); return 0;",
        expect: [""]
    },
    printf_hhd_from_int_promotion => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "printf(\"%hhd\\n\", (signed char)127); return 0;",
        expect: ["127"]
    },
    printf_hd_from_expression => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "printf(\"%hd\\n\", (short)(100+23)); return 0;",
        expect: ["123"]
    },
    printf_lld_from_multiplication => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "printf(\"%lld\\n\", 1000LL*1000LL); return 0;",
        expect: ["1000000"]
    },
    printf_td_array_stride => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "int m[3][4]; ptrdiff_t d=&m[1][0]-&m[0][0]; printf(\"%td\\n\", d); return 0;",
        expect: ["4"]
    },
    printf_lf_precision_three => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "double d=1.0/3.0; printf(\"%.3lf\\n\", d); return 0;",
        expect: ["0.333"]
    },
    printf_lf_width_ten => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "long double ld=3.5L; printf(\"%10.1Lf\\n\", ld); return 0;",
        expect: ["       3.5"]
    },
    printf_llu_octal => {
        includes: ["<stdio.h>", "<stddef.h>", "<inttypes.h>"],
        decls: "",
        body: "unsigned long long v=64ULL; printf(\"%llo\\n\", v); return 0;",
        expect: ["100"]
    },
}
