//! Pointer, array decay, and qualifier semantics — one language rule per test.

use crate::helpers::*;

c_run_cases! {
    array_decays_to_pointer_in_call => {
        includes: ["<stdio.h>"],
        decls: "void show(int *p) { printf(\"%d\\n\", p[1]); }",
        body: "int a[3] = {10,20,30}; show(a); return 0;",
        expect: ["20"]
    },
    pointer_subscript_equivalent_to_deref_add => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[2] = {5,9}; printf(\"%d\\n\", *(a+1)); return 0;",
        expect: ["9"]
    },
    address_of_array_element => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[2] = {1,2}; int *p = &a[1]; printf(\"%d\\n\", *p); return 0;",
        expect: ["2"]
    },
    pointer_difference_counts_elements => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[4] = {0}; printf(\"%d\\n\", (int)(&a[3] - &a[0])); return 0;",
        expect: ["3"]
    },
    const_pointer_cannot_change_pointee => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x = 1; const int *p = &x; printf(\"%d\\n\", *p); return 0;",
        expect: ["1"]
    },
    pointer_to_const_value => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "const int x = 7; int *p = (int*)&x; printf(\"%d\\n\", *p); return 0;",
        expect: ["7"]
    },
    void_pointer_generic_address => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x = 4; void *p = &x; printf(\"%d\\n\", *(int*)p); return 0;",
        expect: ["4"]
    },
    null_pointer_compare => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int *p = 0; printf(\"%d\\n\", p == 0); return 0;",
        expect: ["1"]
    },
    double_pointer_indirection => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x = 3; int *p = &x; int **pp = &p; printf(\"%d\\n\", **pp); return 0;",
        expect: ["3"]
    },
    pointer_arithmetic_on_char_bytes => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char s[] = \"abc\"; char *p = s; printf(\"%c\\n\", *(p+2)); return 0;",
        expect: ["c"]
    },
    array_parameter_is_pointer => {
        includes: ["<stdio.h>"],
        decls: "int len(int *a) { return sizeof(a) / sizeof(a[0]); }",
        body: "int a[5] = {0}; printf(\"%d\\n\", len(a)); return 0;",
        expect: ["2"]
    },
    static_array_size_in_same_scope => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[5] = {0}; printf(\"%d\\n\", (int)(sizeof(a)/sizeof(a[0]))); return 0;",
        expect: ["5"]
    },
    string_literal_has_static_storage => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char *s = \"hi\"; printf(\"%s\\n\", s); return 0;",
        expect: ["hi"]
    },
    char_array_mutable_copy => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char s[] = \"ab\"; s[0] = 'x'; printf(\"%s\\n\", s); return 0;",
        expect: ["xb"]
    },
    restrict_pointer_alias_hint => {
        includes: ["<stdio.h>"],
        decls: "int add(restrict int *a, restrict int *b) { return *a + *b; }",
        body: "int x=2,y=3; printf(\"%d\\n\", add(&x,&y)); return 0;",
        expect: ["5"]
    },
}

c_compile_cases! {
    incomplete_array_parameter => { includes: ["<stdio.h>"], decls: "void f(int a[]);", body: "return 0;" },
    function_pointer_typedef => { includes: ["<stdio.h>"], decls: "typedef int (*op_t)(int,int);", body: "op_t f = 0; return 0;" },
    pointer_to_function => { includes: ["<stdio.h>"], decls: "int g(int x) { return x; }", body: "int (*fp)(int) = g; return fp(1);" },
    volatile_qualified_load => { includes: ["<stdio.h>"], decls: "", body: "volatile int v = 1; return v;" },
    const_top_level_array_param => { includes: ["<stdio.h>"], decls: "void f(const int *a);", body: "return 0;" },
}
