//! Integer literal lexical forms: bases, suffixes, leading zeros.


c_run_cases! {
    decimal_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0); return 0;",
        expect: ["0"]
    },
    decimal_single_digit_nine => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 9); return 0;",
        expect: ["9"]
    },
    decimal_multidigit_two_fifty_six => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 256); return 0;",
        expect: ["256"]
    },
    decimal_million => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 1000000); return 0;",
        expect: ["1000000"]
    },
    decimal_negative_ninety_nine => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", -99); return 0;",
        expect: ["-99"]
    },
    decimal_unary_plus_seventeen => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", +17); return 0;",
        expect: ["17"]
    },
    leading_zero_octal_seven => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 07); return 0;",
        expect: ["7"]
    },
    leading_zero_octal_ten => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 010); return 0;",
        expect: ["8"]
    },
    leading_zero_octal_sixty_four => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0100); return 0;",
        expect: ["64"]
    },
    octal_three => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 03); return 0;",
        expect: ["3"]
    },
    octal_sixty_three => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 077); return 0;",
        expect: ["63"]
    },
    octal_two_fifty_five => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0377); return 0;",
        expect: ["255"]
    },
    octal_one_two_three => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0123); return 0;",
        expect: ["83"]
    },
    octal_with_unsigned_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 0177777u); return 0;",
        expect: ["65535"]
    },
    hex_zero => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0x0); return 0;",
        expect: ["0"]
    },
    hex_single_digit_fifteen => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0xf); return 0;",
        expect: ["15"]
    },
    hex_two_fifty_six => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0x100); return 0;",
        expect: ["256"]
    },
    hex_mixed_case_dead => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0xDeAd); return 0;",
        expect: ["57005"]
    },
    hex_long_digits => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 0x12345678u); return 0;",
        expect: ["305419896"]
    },
    hex_with_long_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%ld\\n\", 0x10L); return 0;",
        expect: ["16"]
    },
    hex_with_unsigned_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 0x100u); return 0;",
        expect: ["256"]
    },
    hex_with_unsigned_long_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 0xFul); return 0;",
        expect: ["15"]
    },
    suffix_lowercase_u => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 18u); return 0;",
        expect: ["18"]
    },
    suffix_uppercase_u => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 19U); return 0;",
        expect: ["19"]
    },
    suffix_lowercase_l => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%ld\\n\", 20l); return 0;",
        expect: ["20"]
    },
    suffix_uppercase_l => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%ld\\n\", 21L); return 0;",
        expect: ["21"]
    },
    suffix_lowercase_ll => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%lld\\n\", 22LL); return 0;",
        expect: ["22"]
    },
    suffix_uppercase_ll => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%lld\\n\", 23LL); return 0;",
        expect: ["23"]
    },
    suffix_ul_lowercase => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 24ul); return 0;",
        expect: ["24"]
    },
    suffix_ul_uppercase => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 25UL); return 0;",
        expect: ["25"]
    },
    suffix_lu_reversed => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 26lu); return 0;",
        expect: ["26"]
    },
    suffix_lu_reversed_upper => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 27LU); return 0;",
        expect: ["27"]
    },
    unsigned_sixteen_bit_max => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 65535u); return 0;",
        expect: ["65535"]
    },
    long_long_ten_billion => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%lld\\n\", 10000000000LL); return 0;",
        expect: ["10000000000"]
    },
    negative_hex_literal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", -0x10); return 0;",
        expect: ["-16"]
    },
    negative_octal_literal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", -012); return 0;",
        expect: ["-10"]
    },
    mixed_base_addition => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0xA + 012); return 0;",
        expect: ["20"]
    },
    octal_times_decimal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 04 * 5); return 0;",
        expect: ["20"]
    },
    hex_shift_by_hex => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 1 << 0x4); return 0;",
        expect: ["16"]
    },
    decimal_thirty_two_seven_sixty_seven => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 32767); return 0;",
        expect: ["32767"]
    },
    decimal_thirty_two_seven_sixty_eight => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 32768); return 0;",
        expect: ["32768"]
    },
    zero_with_unsigned_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 0u); return 0;",
        expect: ["0"]
    },
    zero_with_long_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%ld\\n\", 0L); return 0;",
        expect: ["0"]
    },
    hex_seven_f_f_f => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0x7FFF); return 0;",
        expect: ["32767"]
    },
    octal_four_four_four => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0444); return 0;",
        expect: ["292"]
    },
    decimal_three_hundred_unsigned => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%u\\n\", 300u); return 0;",
        expect: ["300"]
    },
    hex_one_a => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0x1A); return 0;",
        expect: ["26"]
    },
    octal_with_long_suffix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%ld\\n\", 01l); return 0;",
        expect: ["1"]
    },
    hex_cafe => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0xCAFE); return 0;",
        expect: ["51966"]
    },
    decimal_in_subtraction_with_hex => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 100 - 0x20); return 0;",
        expect: ["68"]
    },
    unsigned_long_long_small => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%llu\\n\", 88ull); return 0;",
        expect: ["88"]
    },
    hex_uppercase_prefix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 0X20); return 0;",
        expect: ["32"]
    },
}
