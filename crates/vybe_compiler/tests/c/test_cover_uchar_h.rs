//! uchar.h — UTF character typedefs, literals, macros, and conversion entry points.


c_run_cases! {
    uchar_char8_literal_value => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "char8_t c = u8'a'; printf(\"%d\\n\", (int)c); return 0;", expect: ["97"] },
    uchar_char16_literal_value => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "char16_t c = u'b'; printf(\"%d\\n\", (int)c); return 0;", expect: ["98"] },
    uchar_char32_literal_value => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "char32_t c = U'c'; printf(\"%d\\n\", (int)c); return 0;", expect: ["99"] },
    uchar_sizeof_types_wasm32 => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "printf(\"%d %d %d\\n\", (int)sizeof(char8_t), (int)sizeof(char16_t), (int)sizeof(char32_t)); return 0;", expect: ["1 2 4"] },
    uchar_stdc_utf_macros_defined => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "printf(\"%d %d\\n\", __STDC_UTF_16__, __STDC_UTF_32__); return 0;", expect: ["1 1"] },
    uchar_mbrtoc16_ascii => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "char16_t out[2]; size_t n = mbrtoc16(out, \"Q\", 1, 0); printf(\"%d %d\\n\", (int)n, (int)out[0]); return 0;", expect: ["1 81"] },
    uchar_mbrtoc32_ascii => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "char32_t out[2]; size_t n = mbrtoc32(out, \"R\", 1, 0); printf(\"%d %d\\n\", (int)n, (int)out[0]); return 0;", expect: ["1 82"] },
    uchar_c16rtomb_ascii => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "char b[4]; size_t n = c16rtomb(b, u'S', 0); printf(\"%d %c\\n\", (int)n, b[0]); return 0;", expect: ["1 S"] },
    uchar_c32rtomb_ascii => { includes: ["<stdio.h>", "<uchar.h>"], decls: "", body: "char b[4]; size_t n = c32rtomb(b, U'T', 0); printf(\"%d %c\\n\", (int)n, b[0]); return 0;", expect: ["1 T"] },
}

c_compile_cases! {
    uchar_mbstate_t_declares => { includes: ["<uchar.h>"], decls: "", body: "mbstate_t st; (void)st; return 0;" },
    uchar_utf8_string_initializes_char8_pointer => { includes: ["<uchar.h>"], decls: "", body: "const char8_t *s = u8\"ok\"; return s[0];" },
    uchar_utf16_string_initializes_array => { includes: ["<uchar.h>"], decls: "", body: "const char16_t *s = u\"ok\"; return s[1];" },
    uchar_utf32_string_initializes_array => { includes: ["<uchar.h>"], decls: "", body: "const char32_t *s = U\"ok\"; return s[1];" },
}
