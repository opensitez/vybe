//! Storage duration, linkage, and translation-unit semantics.

use crate::helpers::*;

c_run_cases! {
    extern_global_read => {
        includes: ["<stdio.h>"],
        decls: "extern int g_ext; int g_ext = 7;",
        body: "printf(\"%d\\n\", g_ext); return 0;",
        expect: ["7"]
    },
    static_file_scope_hidden => {
        includes: ["<stdio.h>"],
        decls: "static int hidden = 4; int read_hidden(void){ return hidden; }",
        body: "printf(\"%d\\n\", read_hidden()); return 0;",
        expect: ["4"]
    },
    static_function_local_persists => {
        includes: ["<stdio.h>"],
        decls: "int tick(void){ static int n; return ++n; }",
        body: "printf(\"%d %d\\n\", tick(), tick()); return 0;",
        expect: ["1 2"]
    },
    const_object_internal_linkage => {
        includes: ["<stdio.h>"],
        decls: "const int cval = 12;",
        body: "printf(\"%d\\n\", cval); return 0;",
        expect: ["12"]
    },
    typedef_same_type_identity => {
        includes: ["<stdio.h>"],
        decls: "typedef int myint; typedef myint myint2;",
        body: "myint2 x=8; printf(\"%d\\n\", x); return 0;",
        expect: ["8"]
    },
    struct_tag_scope => {
        includes: ["<stdio.h>"],
        decls: "struct Node { int v; struct Node *next; };",
        body: "struct Node n={5,0}; printf(\"%d\\n\", n.v); return 0;",
        expect: ["5"]
    },
    union_all_members_alias => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; float f; };",
        body: "union U u; u.f=1.0f; printf(\"%d\\n\", u.i != 0); return 0;",
        expect: ["1"]
    },
    enum_unscoped_constants => {
        includes: ["<stdio.h>"],
        decls: "enum Color { RED, GREEN };",
        body: "printf(\"%d\\n\", GREEN); return 0;",
        expect: ["1"]
    },
    auto_storage_default => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "auto int x = 6; printf(\"%d\\n\", x); return 0;",
        expect: ["6"]
    },
    register_hint_still_usable => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "register int r = 3; printf(\"%d\\n\", r); return 0;",
        expect: ["3"]
    },
}

c_compile_cases! {
    linkage_extern_decl_no_init => { includes: ["<stdio.h>"], decls: "extern int g;", body: "return 0;" },
    linkage_static_extern_function => { includes: ["<stdio.h>"], decls: "static void f(void){} static void f(void){}", body: "f(); return 0;" },
    linkage_inline_extern => { includes: ["<stdio.h>"], decls: "inline int add(int a,int b){return a+b;}", body: "return add(1,2);" },
    linkage_thread_local => { includes: ["<stdio.h>"], decls: "_Thread_local int tls;", body: "tls=1; return tls;" },
    linkage_constexpr_like_enum => { includes: ["<stdio.h>"], decls: "enum { BUF = 64 }; char a[BUF];", body: "return sizeof(a);" },
    linkage_nested_struct_scope => { includes: ["<stdio.h>"], decls: "struct Outer { struct Inner { int x; } in; };", body: "struct Outer o={.in={1}}; return o.in.x;" },
    linkage_anonymous_struct_member => { includes: ["<stdio.h>"], decls: "struct S { struct { int x; }; };", body: "struct S s={.x=2}; return s.x;" },
    linkage_bitfield_signed => { includes: ["<stdio.h>"], decls: "struct B { signed int s:4; };", body: "struct B b={-1}; return b.s;" },
    linkage_flexible_array_typedef => { includes: ["<stdlib.h>"], decls: "typedef struct { int n; char tail[]; } Blob;", body: "Blob *b=malloc(sizeof(Blob)+4); free(b); return 0;" },
    linkage_void_fn_param_list => { includes: ["<stdio.h>"], decls: "void f(void){}", body: "f(); return 0;" },
    linkage_old_style_proto => { includes: ["<stdio.h>"], decls: "int f(); int f(int x){return x;}", body: "return f(3);" },
    linkage_kr_style_def => { includes: ["<stdio.h>"], decls: "int f(x) int x; { return x; }", body: "return f(2);" },
}
