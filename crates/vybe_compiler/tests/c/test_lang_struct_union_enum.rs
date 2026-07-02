//! struct, union, enum language rules — one behavior per test.


c_run_cases! {
    struct_member_access_dot => {
        includes: ["<stdio.h>"],
        decls: "struct P { int x; int y; };",
        body: "struct P p = {.x=2,.y=5}; printf(\"%d\\n\", p.x + p.y); return 0;",
        expect: ["7"]
    },
    struct_pointer_arrow_access => {
        includes: ["<stdio.h>"],
        decls: "struct P { int n; };",
        body: "struct P p = {9}; struct P *pp = &p; printf(\"%d\\n\", pp->n); return 0;",
        expect: ["9"]
    },
    designated_initializer_sparse => {
        includes: ["<stdio.h>"],
        decls: "struct S { int a,b,c; };",
        body: "struct S s = {.c = 4}; printf(\"%d %d\\n\", s.a, s.c); return 0;",
        expect: ["0 4"]
    },
    union_read_active_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; char c; };",
        body: "union U u; u.i = 65; printf(\"%d\\n\", u.i); return 0;",
        expect: ["65"]
    },
    enum_named_constants => {
        includes: ["<stdio.h>"],
        decls: "enum Color { RED, GREEN, BLUE };",
        body: "enum Color c = GREEN; printf(\"%d\\n\", c); return 0;",
        expect: ["1"]
    },
    enum_explicit_values => {
        includes: ["<stdio.h>"],
        decls: "enum E { A = 10, B, C = 20 };",
        body: "printf(\"%d %d\\n\", B, C); return 0;",
        expect: ["11 20"]
    },
    typedef_struct_tag => {
        includes: ["<stdio.h>"],
        decls: "typedef struct { int v; } Box;",
        body: "Box b = {3}; printf(\"%d\\n\", b.v); return 0;",
        expect: ["3"]
    },
    nested_struct_access => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{7}}; printf(\"%d\\n\", o.in.n); return 0;",
        expect: ["7"]
    },
    anonymous_struct_within_union => {
        includes: ["<stdio.h>"],
        decls: "union U { struct { int x; } s; int i; };",
        body: "union U u; u.s.x = 2; printf(\"%d\\n\", u.s.x); return 0;",
        expect: ["2"]
    },
    bitfield_read => {
        includes: ["<stdio.h>"],
        decls: "struct Flags { unsigned a:1; unsigned b:3; };",
        body: "struct Flags f = {1,5}; printf(\"%u %u\\n\", f.a, f.b); return 0;",
        expect: ["1 5"]
    },
    flexible_array_member_size => {
        includes: ["<stdio.h>"],
        decls: "struct Buf { int n; char data[]; };",
        body: "struct Buf *b = malloc(sizeof(struct Buf) + 4); b->n = 4; b->data[0]='a'; printf(\"%c\\n\", b->data[0]); free(b); return 0;",
        expect: ["a"]
    },
    offsetof_macro => {
        includes: ["<stdio.h>", "<stddef.h>"],
        decls: "struct S { char c; int n; };",
        body: "printf(\"%d\\n\", (int)offsetof(struct S, n)); return 0;",
        expect: ["4"]
    },
    alignof_type => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(double)); return 0;",
        expect: ["8"]
    },
    compound_literal_struct => {
        includes: ["<stdio.h>"],
        decls: "struct P { int x; };",
        body: "struct P *p = &(struct P){.x=6}; printf(\"%d\\n\", p->x); return 0;",
        expect: ["6"]
    },
    compound_literal_array => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int *p = (int[]){1,2,3}; printf(\"%d\\n\", p[2]); return 0;",
        expect: ["3"]
    },
}

c_compile_cases! {
    struct_forward_declaration => { includes: ["<stdio.h>"], decls: "struct Node; struct Node { struct Node *next; };", body: "return 0;" },
    enum_opaque_compile => { includes: ["<stdio.h>"], decls: "enum E; enum E { X };", body: "return 0;" },
    packed_struct_attribute_compile => { includes: ["<stdio.h>"], decls: "struct __attribute__((packed)) P { char c; int n; };", body: "return 0;" },
}
