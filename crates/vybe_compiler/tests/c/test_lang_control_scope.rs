//! Control flow, scope, and linkage — one language rule per test.


c_run_cases! {
    switch_fallthrough_behavior => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x=0; switch(1){case 1: x++; case 2: x++; break;} printf(\"%d\\n\", x); return 0;",
        expect: ["2"]
    },
    switch_default_case => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "switch(9){case 1: printf(\"1\\n\"); break; default: printf(\"d\\n\"); break;} return 0;",
        expect: ["d"]
    },
    do_while_executes_once => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int i=0; do { i++; } while(0); printf(\"%d\\n\", i); return 0;",
        expect: ["1"]
    },
    while_loop_counts => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int i=0; while(i<3) i++; printf(\"%d\\n\", i); return 0;",
        expect: ["3"]
    },
    for_loop_scope => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "for(int i=0;i<2;i++){} printf(\"%d\\n\", 1); return 0;",
        expect: ["1"]
    },
    break_exits_loop => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "for(int i=0;i<10;i++){ if(i==2) break; } printf(\"%d\\n\", 2); return 0;",
        expect: ["2"]
    },
    continue_skips_iteration => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int s=0; for(int i=0;i<3;i++){ if(i==1) continue; s+=i; } printf(\"%d\\n\", s); return 0;",
        expect: ["2"]
    },
    goto_forward_jump => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "goto L; L: printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
    ternary_operator => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a=1,b=2; printf(\"%d\\n\", a>b?a:b); return 0;",
        expect: ["2"]
    },
    comma_operator_value => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x=(1,2,3); printf(\"%d\\n\", x); return 0;",
        expect: ["3"]
    },
    static_local_retains => {
        includes: ["<stdio.h>"],
        decls: "int bump(){ static int c=0; c++; return c; }",
        body: "printf(\"%d %d\\n\", bump(), bump()); return 0;",
        expect: ["1 2"]
    },
    extern_global_access => {
        includes: ["<stdio.h>"],
        decls: "int g_val = 5;",
        body: "printf(\"%d\\n\", g_val); return 0;",
        expect: ["5"]
    },
    block_scope_shadow => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x=1; { int x=2; } printf(\"%d\\n\", x); return 0;",
        expect: ["1"]
    },
    if_else_chain => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int v=2; if(v==1) printf(\"1\\n\"); else if(v==2) printf(\"2\\n\"); else printf(\"x\\n\"); return 0;",
        expect: ["2"]
    },
    logical_and_short_circuit => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x=0; if(0 && (x=1)) {} printf(\"%d\\n\", x); return 0;",
        expect: ["0"]
    },
}

c_compile_cases! {
    internal_static_function => { includes: ["<stdio.h>"], decls: "static int helper(void){return 1;}", body: "return helper();" },
    extern_declaration => { includes: ["<stdio.h>"], decls: "extern int ext_val;", body: "return 0;" },
    nested_block_compile => { includes: ["<stdio.h>"], decls: "", body: "{ { int x=1; } } return 0;" },
    switch_nested_compile => { includes: ["<stdio.h>"], decls: "", body: "switch(1){case 1: switch(2){case 2: break;} break;} return 0;" },
    labeled_statement_compile => { includes: ["<stdio.h>"], decls: "", body: "L: return 0;" },
}
