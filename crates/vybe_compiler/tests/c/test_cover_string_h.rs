//! string.h — one distinct API per test (breadth, not variants).

use crate::helpers::*;

c_run_cases! {
    memccpy_copies_until_char => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "char dst[8]; char *end = memccpy(dst, \"hello\", 'l', 5); *end = '\\0'; printf(\"%s\\n\", dst); return 0;",
        expect: ["hel"]
    },
    strndup_truncates => {
        includes: ["<stdio.h>", "<string.h>", "<stdlib.h>"],
        decls: "",
        body: "char *s = strndup(\"abcdef\", 3); printf(\"%s\\n\", s); free(s); return 0;",
        expect: ["abc"]
    },
    memchr_finds_byte => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "char *p = memchr(\"abcd\", 'c', 4); printf(\"%s\\n\", p); return 0;",
        expect: ["cd"]
    },
    memrchr_finds_last_byte => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "char *p = memrchr(\"abca\", 'a', 4); printf(\"%c\\n\", *p); return 0;",
        expect: ["a"]
    },
    memcpy_copies_bytes => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "char dst[4]; memcpy(dst, \"xyz\", 4); printf(\"%s\\n\", dst); return 0;",
        expect: ["xyz"]
    },
    memmove_overlapping => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "char s[] = \"abcde\"; memmove(s+1, s, 4); printf(\"%s\\n\", s); return 0;",
        expect: ["aabbc"]
    },
    memset_fills => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "char b[3]; memset(b, 'x', 3); b[2]='\\0'; printf(\"%s\\n\", b); return 0;",
        expect: ["xx"]
    },
    memcmp_lexicographic => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "printf(\"%d\\n\", memcmp(\"abc\", \"abd\", 3)); return 0;",
        expect: ["-1"]
    },
    strlen_counts => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)strlen(\"go\")); return 0;",
        expect: ["2"]
    },
    strcpy_copies => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "char d[8]; strcpy(d, \"vy\"); printf(\"%s\\n\", d); return 0;",
        expect: ["vy"]
    },
    strcat_appends => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "char d[8] = \"a\"; strcat(d, \"b\"); printf(\"%s\\n\", d); return 0;",
        expect: ["ab"]
    },
    strcmp_equal => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "printf(\"%d\\n\", strcmp(\"ab\", \"ab\")); return 0;",
        expect: ["0"]
    },
    strchr_finds_char => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "printf(\"%s\\n\", strchr(\"hello\", 'l')); return 0;",
        expect: ["llo"]
    },
    strrchr_finds_last => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "",
        body: "printf(\"%s\\n\", strrchr(\"hello\", 'l')); return 0;",
        expect: ["lo"]
    },
    strerror_returns_message => {
        includes: ["<stdio.h>", "<string.h>", "<errno.h>"],
        decls: "",
        body: "printf(\"%d\\n\", strerror(EINVAL)[0] != '\\0'); return 0;",
        expect: ["1"]
    },
}

c_compile_cases! {
    strdup_compile => { includes: ["<string.h>", "<stdlib.h>"], decls: "", body: "char *s = strdup(\"x\"); free(s); return 0;" },
    strncpy_compile => { includes: ["<string.h>"], decls: "", body: "char d[4]; strncpy(d, \"abc\", 3); return 0;" },
    strncat_compile => { includes: ["<string.h>"], decls: "", body: "char d[8]=\"a\"; strncat(d,\"b\",1); return 0;" },
    strncmp_compile => { includes: ["<string.h>"], decls: "", body: "return strncmp(\"a\",\"b\",1);" },
    strpbrk_compile => { includes: ["<string.h>"], decls: "", body: "return strpbrk(\"ab\",\"b\") != 0;" },
    strspn_compile => { includes: ["<string.h>"], decls: "", body: "return (int)strspn(\"abc\",\"ab\");" },
    strcspn_compile => { includes: ["<string.h>"], decls: "", body: "return (int)strcspn(\"abc\",\"b\");" },
    strstr_compile => { includes: ["<string.h>"], decls: "", body: "return strstr(\"ab\",\"b\") != 0;" },
    strtok_compile => { includes: ["<string.h>"], decls: "", body: "char s[]=\"a:b\"; strtok(s,\":\"); return 0;" },
}
