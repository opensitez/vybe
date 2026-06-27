//! Arrays, VLAs, and memory layout — one language rule per test.

use crate::helpers::*;

c_run_cases! {
    vla_size_runtime => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; int a[n]; a[2]=9; printf(\"%d\\n\", a[2]); return 0;",
        expect: ["9"]
    },
    vla_multidimensional => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int r=2,c=2; int m[r][c]; m[1][1]=4; printf(\"%d\\n\", m[1][1]); return 0;",
        expect: ["4"]
    },
    array_to_pointer_decay_assignment => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[2]={1,2}; int *p=a; printf(\"%d\\n\", p[1]); return 0;",
        expect: ["2"]
    },
    multidim_row_major_index => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int m[2][3]={{1,2,3},{4,5,6}}; printf(\"%d\\n\", m[1][2]); return 0;",
        expect: ["6"]
    },
    string_literal_concatenation => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%s\\n\", \"hel\" \"lo\"); return 0;",
        expect: ["hello"]
    },
    wide_string_literal => {
        includes: ["<stdio.h>", "<wchar.h>"],
        decls: "",
        body: "wchar_t *s = L\"w\"; printf(\"%lc\\n\", s[0]); return 0;",
        expect: ["w"]
    },
    compound_literal_address_persists_in_scope => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int *p = (int[]){1,2,3}; printf(\"%d\\n\", p[2]); return 0;",
        expect: ["3"]
    },
    struct_padding_read => {
        includes: ["<stdio.h>"],
        decls: "struct S { char c; int n; };",
        body: "struct S s = {.c='a', .n=3}; printf(\"%d\\n\", s.n); return 0;",
        expect: ["3"]
    },
    union_size_reuse => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; double d; };",
        body: "union U u; u.i=42; printf(\"%d\\n\", u.i); return 0;",
        expect: ["42"]
    },
    enum_underlying_arithmetic => {
        includes: ["<stdio.h>"],
        decls: "enum E { A=1, B=2 };",
        body: "enum E e = A; printf(\"%d\\n\", e + B); return 0;",
        expect: ["3"]
    },
    pointer_to_array_type => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[2][3]={{1,2,3},{4,5,6}}; int (*p)[3]=a; printf(\"%d\\n\", p[1][2]); return 0;",
        expect: ["6"]
    },
    stack_array_zero_init => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[3]={0}; printf(\"%d\\n\", a[2]); return 0;",
        expect: ["0"]
    },
    memcpy_struct_assignment => {
        includes: ["<stdio.h>", "<string.h>"],
        decls: "struct P { int x; };",
        body: "struct P a={1},b; b=a; printf(\"%d\\n\", b.x); return 0;",
        expect: ["1"]
    },
}

c_compile_cases! {
    flexible_array_in_struct => { includes: ["<stdlib.h>"], decls: "struct B { int n; char d[]; };", body: "struct B *b=malloc(sizeof(struct B)+4); free(b); return 0;" },
    vla_in_loop_compile => { includes: ["<stdio.h>"], decls: "", body: "for(int i=1;i<2;i++){ int a[i]; (void)a; } return 0;" },
    incomplete_array_extern => { includes: ["<stdio.h>"], decls: "extern int arr[];", body: "return 0;" },
    zero_length_array_extension => { includes: ["<stdio.h>"], decls: "struct Z { int n; int a[0]; };", body: "return sizeof(struct Z);" },
}
