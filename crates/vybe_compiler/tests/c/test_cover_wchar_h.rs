//! wchar.h — one wide-character API per test.

use crate::helpers::*;

c_run_cases! {
    wcslen_counts => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "wchar_t s[] = L\"ab\"; printf(\"%d\\n\", (int)wcslen(s)); return 0;", expect: ["2"] },
    wcscpy_copies => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "wchar_t d[4]; wcscpy(d, L\"go\"); printf(\"%lc\\n\", d[0]); return 0;", expect: ["g"] },
    wcscmp_equal => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "printf(\"%d\\n\", wcscmp(L\"a\", L\"a\")); return 0;", expect: ["0"] },
    wcschr_finds => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "wchar_t *p = wcschr(L\"abc\", L'b'); printf(\"%lc\\n\", *p); return 0;", expect: ["b"] },
    wcsrchr_finds_last => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "wchar_t *p = wcsrchr(L\"abcb\", L'b'); printf(\"%lc\\n\", *p); return 0;", expect: ["b"] },
    wcsncmp_prefix => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "printf(\"%d\\n\", wcsncmp(L\"abc\", L\"abd\", 2)); return 0;", expect: ["0"] },
    wcsncpy_truncates => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "wchar_t d[4]; wcsncpy(d, L\"abcdef\", 3); d[3]=L'\\0'; printf(\"%d\\n\", (int)wcslen(d)); return 0;", expect: ["3"] },
    wcscat_appends => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "wchar_t d[8]=L\"a\"; wcscat(d, L\"b\"); printf(\"%lc\\n\", d[1]); return 0;", expect: ["b"] },
    btowc_ascii => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "printf(\"%lc\\n\", btowc('A')); return 0;", expect: ["A"] },
    wctob_ascii => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "printf(\"%c\\n\", (char)wctob(L'Z')); return 0;", expect: ["Z"] },
}

c_compile_cases! {
    mbsrtowcs_compile => { includes: ["<wchar.h>", "<stdlib.h>"], decls: "", body: "const char *src = \"a\"; wchar_t dst[4]; mbsrtowcs(dst, &src, 4, 0); return 0;" },
    wcsrtombs_compile => { includes: ["<wchar.h>", "<stdlib.h>"], decls: "", body: "const wchar_t *src = L\"a\"; char dst[4]; wcsrtombs(dst, &src, 4, 0); return 0;" },
    mbstowcs_compile => { includes: ["<stdlib.h>"], decls: "", body: "wchar_t w[4]; mbstowcs(w, \"a\", 4); return 0;" },
    wcstombs_compile => { includes: ["<stdlib.h>"], decls: "", body: "char b[4]; wcstombs(b, L\"a\", 4); return 0;" },
    wmemchr_compile => { includes: ["<wchar.h>"], decls: "", body: "return wmemchr(L\"abc\", L'b', 3) != 0;" },
    wmemcmp_compile => { includes: ["<wchar.h>"], decls: "", body: "return wmemcmp(L\"a\", L\"b\", 1);" },
    wmemcpy_compile => { includes: ["<wchar.h>"], decls: "", body: "wchar_t d[2]; wmemcpy(d, L\"a\", 2); return 0;" },
    wmemset_compile => { includes: ["<wchar.h>"], decls: "", body: "wchar_t d[2]; wmemset(d, L'x', 2); return 0;" },
}
