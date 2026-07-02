//! Operators, promotions, and casts — one rule per test.


c_run_cases! {
    integer_promotion_in_add => { includes: ["<stdio.h>"], decls: "", body: "char a=1,b=2; printf(\"%d\\n\", a+b); return 0;", expect: ["3"] },
    usual_arithmetic_conversions => { includes: ["<stdio.h>"], decls: "", body: "int i=2; double d=2.5; printf(\"%.1f\\n\", i+d); return 0;", expect: ["4.5"] },
    float_division => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%.1f\\n\", 5.0/2.0); return 0;", expect: ["2.5"] },
    integer_division_truncates => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 5/2); return 0;", expect: ["2"] },
    modulo_operator => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 10%3); return 0;", expect: ["1"] },
    left_shift => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 1<<4); return 0;", expect: ["16"] },
    right_shift => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 16>>2); return 0;", expect: ["4"] },
    bitwise_and => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 6&3); return 0;", expect: ["2"] },
    bitwise_or => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 4|1); return 0;", expect: ["5"] },
    bitwise_xor => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 10^12); return 0;", expect: ["6"] },
    unary_bitwise_not => { includes: ["<stdio.h>"], decls: "", body: "unsigned char c=0; printf(\"%u\\n\", (unsigned)(~c & 0xFF)); return 0;", expect: ["255"] },
    logical_not => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", !0); return 0;", expect: ["1"] },
    cast_int_to_char => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%c\\n\", (char)66); return 0;", expect: ["B"] },
    cast_double_to_int_trunc => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", (int)3.9); return 0;", expect: ["3"] },
    implicit_bool_in_condition => { includes: ["<stdio.h>"], decls: "", body: "int ok=1; if(ok) printf(\"y\\n\"); return 0;", expect: ["y"] },
    sizeof_type_operand => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", (int)sizeof(int)); return 0;", expect: ["4"] },
    sizeof_expression_no_eval => { includes: ["<stdio.h>"], decls: "", body: "int x=0; printf(\"%d\\n\", (int)sizeof(x++)); return 0;", expect: ["4"] },
    pointer_cast_void_roundtrip => { includes: ["<stdio.h>"], decls: "", body: "int x=8; void *p=&x; printf(\"%d\\n\", *(int*)p); return 0;", expect: ["8"] },
    const_cast_discards_qualifier => { includes: ["<stdio.h>"], decls: "", body: "const int c=3; int *p=(int*)&c; printf(\"%d\\n\", *p); return 0;", expect: ["3"] },
    compound_assignment_mul => { includes: ["<stdio.h>"], decls: "", body: "int n=2; n*=3; printf(\"%d\\n\", n); return 0;", expect: ["6"] },
    compound_assignment_shift => { includes: ["<stdio.h>"], decls: "", body: "int n=1; n<<=2; printf(\"%d\\n\", n); return 0;", expect: ["4"] },
    increment_prefix => { includes: ["<stdio.h>"], decls: "", body: "int n=1; printf(\"%d\\n\", ++n); return 0;", expect: ["2"] },
    increment_postfix => { includes: ["<stdio.h>"], decls: "", body: "int n=1; int m=n++; printf(\"%d %d\\n\", m, n); return 0;", expect: ["1 2"] },
    relational_chained => { includes: ["<stdio.h>"], decls: "", body: "printf(\"%d\\n\", 1<2 && 2<3); return 0;", expect: ["1"] },
    equality_struct_bitwise => { includes: ["<stdio.h>"], decls: "struct P{int x;};", body: "struct P a={1},b={1}; printf(\"%d\\n\", a.x==b.x); return 0;", expect: ["1"] },
}

c_compile_cases! {
    cast_incomplete_array => { includes: ["<stdio.h>"], decls: "", body: "int *p = (int*)(void*)0; return 0;" },
    generic_selection => { includes: ["<stdio.h>"], decls: "#define TYPE(x) _Generic((x), int: 1, default: 0)", body: "return TYPE(0);" },
    typeof_unary_compile => { includes: ["<stdio.h>"], decls: "", body: "int x=1; typeof(x) y=2; return y;" },
}
