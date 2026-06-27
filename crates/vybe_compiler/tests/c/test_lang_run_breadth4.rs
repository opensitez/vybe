//! Language runtime batch 4 — distinct rules not covered in breadth 1–3.

use crate::helpers::*;

c_run_cases! {
    sizeof_string_literal_includes_nul => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%zu\\n\", sizeof \"ab\"); return 0;", expect: ["3"] },
    sizeof_array_not_pointer => { includes: ["<stdio.h>"], decls: "", body: "int a[4]; printf(\"%zu\\n\", sizeof a); return 0;", expect: ["16"] },
    array_decay_only_in_param => {
        includes: ["<stdio.h>"],
        decls: "int a[4]; int sz(void){ return (int)sizeof a; }",
        body: "printf(\"%d\\n\", sz()); return 0;",
        expect: ["16"]
    },
    compound_literal_address => { includes: ["<stdio.h>"], decls: "", body: "int *p = (int[]){1,2,3}; printf(\"%d\\n\", p[2]); return 0;", expect: ["3"] },
    designated_struct_init => { includes: ["<stdio.h>"], decls: "struct S{int a,b,c;};", body: "struct S s={.b=5}; printf(\"%d\\n\", s.b); return 0;", expect: ["5"] },
    designated_array_init => { includes: ["<stdio.h>"], decls: "", body: "int a[5]={[2]=9}; printf(\"%d\\n\", a[2]); return 0;", expect: ["9"] },
    static_local_retains => {
        includes: ["<stdio.h>"],
        decls: "int bump(void){ static int n; return ++n; }",
        body: "printf(\"%d %d\\n\", bump(), bump()); return 0;",
        expect: ["1 2"]
    },
    incomplete_type_pointer_size => { includes: ["<stdio.h>"], decls: "struct X;", body: "struct X *p=0; printf(\"%d\\n\", p==0); return 0;", expect: ["1"] },
    void_pointer_any_ptr => { includes: ["<stdio.h>"], decls: "", body: "int x=3; void *p=&x; printf(\"%d\\n\", *(int*)p); return 0;", expect: ["3"] },
    null_macro_zero => { includes: ["<stdio.h>", "<stddef.h>"], decls: "", body: "printf(\"%d\\n\", NULL==0); return 0;", expect: ["1"] },
    offsetof_member => { includes: ["<stdio.h>", "<stddef.h>"], decls: "struct S{char c; int n;};", body: "printf(\"%zu\\n\", offsetof(struct S, n)); return 0;", expect: ["4"] },
    alignof_int => { includes: ["<stdio.h>", "<stdalign.h>"], decls: "", body: "printf(\"%zu\\n\", alignof(int)); return 0;", expect: ["4"] },
    multidim_row_major => { includes: ["<stdio.h>"], decls: "", body: "int m[2][3]={{1,2,3},{4,5,6}}; printf(\"%d\\n\", m[1][2]); return 0;", expect: ["6"] },
    pointer_to_multidim => { includes: ["<stdio.h>"], decls: "", body: "int m[2][2]={{1,2},{3,4}}; int (*p)[2]=m; printf(\"%d\\n\", p[1][0]); return 0;", expect: ["3"] },
    function_pointer_call => {
        includes: ["<stdio.h>"],
        decls: "int twice(int x){return x*2;}",
        body: "int (*fp)(int)=twice; printf(\"%d\\n\", fp(4)); return 0;",
        expect: ["8"]
    },
    default_argument_promotion => {
        includes: ["<stdio.h>", "<stdarg.h>"],
        decls: "int sum2(int a, int b){return a+b;}",
        body: "printf(\"%d\\n\", sum2((char)1,(char)2)); return 0;",
        expect: ["3"]
    },
    switch_on_char => { includes: ["<stdio.h>"], decls: "", body: "switch('b'){case 'b': printf(\"ok\\n\"); break; default: printf(\"no\\n\");} return 0;", expect: ["ok"] },
    switch_default_only => { includes: ["<stdio.h>"], decls: "", body: "switch(9){default: printf(\"d\\n\"); break;} return 0;", expect: ["d"] },
    empty_for_body => { includes: ["<stdio.h>"], decls: "", body: "int i=0; for(;i<1;i++){} printf(\"%d\\n\", i); return 0;", expect: ["1"] },
    goto_forward_decl => { includes: ["<stdio.h>"], decls: "", body: "goto L; L: printf(\"1\\n\"); return 0;", expect: ["1"] },
    comma_sequence_value => { includes: ["<stdio.h>"], decls: "", body: "int x=(1,2,3); printf(\"%d\\n\", x); return 0;", expect: ["3"] },
    cast_removes_const_volatile => { includes: ["<stdio.h>"], decls: "", body: "const volatile int x=2; int y=(int)x; printf(\"%d\\n\", y); return 0;", expect: ["2"] },
    bitfield_unsigned_wrap => { includes: ["<stdio.h>"], decls: "struct B{unsigned x:2;};", body: "struct B b={3}; b.x++; printf(\"%u\\n\", b.x); return 0;", expect: ["0"] },
    anonymous_union_in_struct => {
        includes: ["<stdio.h>"],
        decls: "struct S{ union { int i; char c; }; };",
        body: "struct S s; s.i=65; printf(\"%c\\n\", s.c); return 0;",
        expect: ["A"]
    },
    enum_underlying_compare => { includes: ["<stdio.h>"], decls: "enum E{A=10,B};", body: "enum E e=B; printf(\"%d\\n\", e==11); return 0;", expect: ["1"] },
    string_literal_concat => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%s\\n\", \"hel\" \"lo\"); return 0;", expect: ["hello"] },
    wide_string_prefix => { includes: ["<stdio.h>", "<wchar.h>"], decls: "", body: "wchar_t w=L'z'; wprintf(L\"%lc\\n\", w); return 0;", expect: ["z"] },
    hex_escape_in_char => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", '\\x41'); return 0;", expect: ["65"] },
    octal_escape_in_char => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", '\\101'); return 0;", expect: ["65"] },
    multichar_constant => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 'ab' != 0); return 0;", expect: ["1"] },
}
