//! Core C language semantics batch — one distinct rule per test.

c_run_cases! {
    struct_return_by_value => {
        includes: ["<stdio.h>"],
        decls: "struct P { int x; }; struct P make(void){ struct P p={3}; return p; }",
        body: "printf(\"%d\\n\", make().x); return 0;",
        expect: ["3"]
    },
    function_returns_pointer_to_static => {
        includes: ["<stdio.h>"],
        decls: "int *id(void){ static int v=9; return &v; }",
        body: "printf(\"%d\\n\", *id()); return 0;",
        expect: ["9"]
    },
    global_aggregate_initializer => {
        includes: ["<stdio.h>"],
        decls: "int a[3] = {1,2,3};",
        body: "printf(\"%d\\n\", a[2]); return 0;",
        expect: ["3"]
    },
    local_aggregate_partial_init => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[4] = {1,2}; printf(\"%d\\n\", a[3]); return 0;",
        expect: ["0"]
    },
    enum_implicit_increment => {
        includes: ["<stdio.h>"],
        decls: "enum E { X, Y, Z };",
        body: "enum E e = Y; printf(\"%d\\n\", e); return 0;",
        expect: ["1"]
    },
    switch_on_enum => {
        includes: ["<stdio.h>"],
        decls: "enum E { A, B };",
        body: "enum E e=B; switch(e){case B: printf(\"b\\n\"); break; default: printf(\"x\\n\");} return 0;",
        expect: ["b"]
    },
    typedef_function_pointer => {
        includes: ["<stdio.h>"],
        decls: "typedef int (*binop_t)(int,int); int add(int a,int b){return a+b;}",
        body: "binop_t f=add; printf(\"%d\\n\", f(2,4)); return 0;",
        expect: ["6"]
    },
    const_pointer_to_const_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "const int x=5; const int *p=&x; printf(\"%d\\n\", *p); return 0;",
        expect: ["5"]
    },
    volatile_load_store => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "volatile int v=2; v=3; printf(\"%d\\n\", v); return 0;",
        expect: ["3"]
    },
    inline_static_in_file => {
        includes: ["<stdio.h>"],
        decls: "static inline int twice(int x){return x*2;}",
        body: "printf(\"%d\\n\", twice(6)); return 0;",
        expect: ["12"]
    },
    empty_translation_unit_link => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
    char_signedness_behavior => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c=127; printf(\"%d\\n\", c>0); return 0;",
        expect: ["1"]
    },
    unsigned_wrap => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "unsigned char u=255; u++; printf(\"%u\\n\", (unsigned)u); return 0;",
        expect: ["0"]
    },
    signed_left_shift => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", 1<<3); return 0;",
        expect: ["8"]
    },
    pointer_comparison_equal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x=1; int *a=&x,*b=&x; printf(\"%d\\n\", a==b); return 0;",
        expect: ["1"]
    },
    void_pointer_generic => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x=2; void *p=&x; printf(\"%d\\n\", *(int*)p); return 0;",
        expect: ["2"]
    },
    struct_bitfield_store => {
        includes: ["<stdio.h>"],
        decls: "struct F { unsigned a:3; unsigned b:5; };",
        body: "struct F f={.a=3,.b=7}; printf(\"%u %u\\n\", f.a, f.b); return 0;",
        expect: ["3 7"]
    },
    union_type_punning_read => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; unsigned char b[4]; };",
        body: "union U u; u.i=1; printf(\"%u\\n\", u.b[0]); return 0;",
        expect: ["1"]
    },
    array_parameter_decay => {
        includes: ["<stdio.h>"],
        decls: "int sum(int *a, int n){ int t=0; for(int i=0;i<n;i++) t+=a[i]; return t; }",
        body: "int a[]={1,2,3}; printf(\"%d\\n\", sum(a,3)); return 0;",
        expect: ["6"]
    },
    file_scope_static_internal => {
        includes: ["<stdio.h>"],
        decls: "static int hidden = 8;",
        body: "printf(\"%d\\n\", hidden); return 0;",
        expect: ["8"]
    },
}

c_compile_cases! {
    incomplete_enum_compile => { includes: ["<stdio.h>"], decls: "enum E; enum E { V };", body: "return V;" },
    struct_self_pointer => { includes: ["<stdio.h>"], decls: "struct N { struct N *next; };", body: "struct N n={0}; return 0;" },
    function_no_prototype_legacy => { includes: ["<stdio.h>"], decls: "int legacy(); int legacy(){return 1;}", body: "return legacy();" },
    void_expr_compile => { includes: ["<stdio.h>"], decls: "void noop(void){}", body: "noop(); return 0;" },
    compound_literal_const => { includes: ["<stdio.h>"], decls: "", body: "int *p=(int[]){1,2}; return p[0];" },
    alignof_expression => { includes: ["<stdalign.h>"], decls: "", body: "return (int)alignof(double);" },
    max_align_t_size => { includes: ["<stddef.h>"], decls: "", body: "return (int)sizeof(max_align_t);" },
    nullptr_constant => { includes: ["<stddef.h>"], decls: "", body: "return NULL == 0;" },
}
