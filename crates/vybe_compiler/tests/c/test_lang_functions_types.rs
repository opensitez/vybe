//! Function types, prototypes, and varargs — one rule per test.


c_run_cases! {
    prototype_before_definition => {
        includes: ["<stdio.h>"],
        decls: "int add(int,int); int add(int a,int b){return a+b;}",
        body: "printf(\"%d\\n\", add(2,3)); return 0;",
        expect: ["5"]
    },
    void_parameter_list => {
        includes: ["<stdio.h>"],
        decls: "void noop(void) {}",
        body: "noop(); printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
    static_inline_function => {
        includes: ["<stdio.h>"],
        decls: "static inline int dbl(int x){return x*2;}",
        body: "printf(\"%d\\n\", dbl(4)); return 0;",
        expect: ["8"]
    },
    function_pointer_call => {
        includes: ["<stdio.h>"],
        decls: "int inc(int x){return x+1;}",
        body: "int (*fp)(int)=inc; printf(\"%d\\n\", fp(2)); return 0;",
        expect: ["3"]
    },
    callback_via_pointer => {
        includes: ["<stdio.h>"],
        decls: "int apply(int x,int(*f)(int)){return f(x);} int sq(int x){return x*x;}",
        body: "printf(\"%d\\n\", apply(4,sq)); return 0;",
        expect: ["16"]
    },
    struct_with_function_pointer_field => {
        includes: ["<stdio.h>"],
        decls: "typedef struct { int (*op)(int); } VTable; int neg(int x){return -x;}",
        body: "VTable vt={neg}; printf(\"%d\\n\", vt.op(5)); return 0;",
        expect: ["-5"]
    },
    varargs_sum_three => {
        includes: ["<stdio.h>", "<stdarg.h>"],
        decls: "int sum3(int n,...){ va_list ap; va_start(ap,n); int t=0; for(int i=0;i<n;i++) t+=va_arg(ap,int); va_end(ap); return t; }",
        body: "printf(\"%d\\n\", sum3(3,1,2,3)); return 0;",
        expect: ["6"]
    },
    va_copy_macro => {
        includes: ["<stdio.h>", "<stdarg.h>"],
        decls: "int first(int n,...){ va_list ap,ap2; va_start(ap,n); va_copy(ap2,ap); int a=va_arg(ap,int); int b=va_arg(ap2,int); va_end(ap); va_end(ap2); return a+b; }",
        body: "printf(\"%d\\n\", first(2,4,5)); return 0;",
        expect: ["9"]
    },
    noreturn_attribute_compile_use => {
        includes: ["<stdio.h>", "<stdnoreturn.h>"],
        decls: "_Noreturn void stop(void){ for(;;){} }",
        body: "printf(\"1\\n\"); return 0;",
        expect: ["1"]
    },
}

c_compile_cases! {
    k_and_r_style_unspecified_params => { includes: ["<stdio.h>"], decls: "int legacy();", body: "return 0;" },
    array_to_pointer_param_decay => { includes: ["<stdio.h>"], decls: "void take(int *a){}", body: "int x[2]={1,2}; take(x); return 0;" },
    const_param_pointer => { includes: ["<stdio.h>"], decls: "void ro(const int *p){}", body: "int x=1; ro(&x); return 0;" },
    nested_function_statement_compile => { includes: ["<stdio.h>"], decls: "", body: "return 0;" },
    attribute_unused_compile => { includes: ["<stdio.h>"], decls: "__attribute__((unused)) static int u = 0;", body: "return 0;" },
}
