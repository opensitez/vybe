//! Preprocessor and macro semantics — one distinct directive/pattern per test.

use crate::helpers::*;

c_compile_cases! {
    define_object_macro => { includes: ["<stdio.h>"], decls: "#define N 5", body: "return N;" },
    define_function_macro => { includes: ["<stdio.h>"], decls: "#define SQ(x) ((x)*(x))", body: "return SQ(3);" },
    define_stringify => { includes: ["<stdio.h>"], decls: "#define STR(x) #x", body: "return 0;" },
    define_token_paste => { includes: ["<stdio.h>"], decls: "#define CONCAT(a,b) a##b\nint xy = 1;", body: "return CONCAT(x,y);" },
    ifdef_defined => { includes: ["<stdio.h>"], decls: "#define F\n#ifdef F\nint a=1;\n#endif", body: "return a;" },
    ifndef_guard => { includes: ["<stdio.h>"], decls: "#ifndef G\n#define G\nint b=2;\n#endif", body: "return b;" },
    if_expression => { includes: ["<stdio.h>"], decls: "#if 1+1==2\nint c=3;\n#endif", body: "return c;" },
    elif_branch => { includes: ["<stdio.h>"], decls: "#if 0\nint d=1;\n#elif 1\nint d=2;\n#endif", body: "return d;" },
    else_branch => { includes: ["<stdio.h>"], decls: "#if 0\nint e=1;\n#else\nint e=2;\n#endif", body: "return e;" },
    undef_macro => { includes: ["<stdio.h>"], decls: "#define Z 1\n#undef Z\nint f=2;", body: "return f;" },
    include_guard_pattern => { includes: ["<stdio.h>"], decls: "#ifndef H\n#define H\n#endif", body: "return 0;" },
    macro_line_splice => { includes: ["<stdio.h>"], decls: "#define M 1 \\\n+ 2", body: "return M;" },
    defined_operator => { includes: ["<stdio.h>"], decls: "#if defined(__STDC__)\nint g=1;\n#endif", body: "return g;" },
    pragma_once_compile => { includes: ["<stdio.h>"], decls: "#pragma once", body: "return 0;" },
    macro_variadic => { includes: ["<stdio.h>"], decls: "#define LOG(fmt, ...) printf(fmt, __VA_ARGS__)", body: "LOG(\"%d\\n\",1); return 0;" },
    macro_select_gnu => { includes: ["<stdio.h>"], decls: "#define VAL(x) _Generic((x), int: 1, default: 0)", body: "return VAL(0);" },
    static_assert_msg => { includes: ["<assert.h>"], decls: "_Static_assert(sizeof(int)>=4, \"int\");", body: "return 0;" },
    alignas_macro => { includes: ["<stdalign.h>"], decls: "alignas(8) int x;", body: "return x;" },
    noreturn_stddef => { includes: ["<stdnoreturn.h>"], decls: "_Noreturn void halt(void); void halt(void){for(;;){}}", body: "return 0;" },
    thread_local_macro => { includes: ["<stdio.h>"], decls: "_Thread_local int tls;", body: "tls=1; return tls;" },
}

c_run_cases! {
    macro_expands_in_printf => {
        includes: ["<stdio.h>"],
        decls: "#define MSG \"ok\"",
        body: "printf(\"%s\\n\", MSG); return 0;",
        expect: ["ok"]
    },
    conditional_compilation_selects_code => {
        includes: ["<stdio.h>"],
        decls: "#define USE_A 1\n#if USE_A\n#define VAL 7\n#else\n#define VAL 0\n#endif",
        body: "printf(\"%d\\n\", VAL); return 0;",
        expect: ["7"]
    },
}
