//! stdlib.h — one distinct API per test.

c_run_cases! {
    abs_int => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%d\\n\", abs(-9)); return 0;", expect: ["9"] },
    labs_long => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%ld\\n\", labs(-5L)); return 0;", expect: ["5"] },
    llabs_longlong => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%lld\\n\", llabs(-7LL)); return 0;", expect: ["7"] },
    atoi_decimal => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%d\\n\", atoi(\"42\")); return 0;", expect: ["42"] },
    atol_long => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%ld\\n\", atol(\"99\")); return 0;", expect: ["99"] },
    atoll_longlong => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%lld\\n\", atoll(\"1000\")); return 0;", expect: ["1000"] },
    atof_float => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%.1f\\n\", atof(\"3.5\")); return 0;", expect: ["3.5"] },
    strtod_parse => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "char *end; double v = strtod(\"1.25\", &end); printf(\"%.2f\\n\", v); return 0;", expect: ["1.25"] },
    strtof_parse => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%.1f\\n\", strtof(\"2.5\", 0)); return 0;", expect: ["2.5"] },
    strtol_base16 => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%ld\\n\", strtol(\"ff\", 0, 16)); return 0;", expect: ["255"] },
    strtoul_unsigned => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%lu\\n\", strtoul(\"10\", 0, 10)); return 0;", expect: ["10"] },
    strtoll_signed => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%lld\\n\", strtoll(\"-8\", 0, 10)); return 0;", expect: ["-8"] },
    strtoull_unsigned => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%llu\\n\", strtoull(\"9\", 0, 10)); return 0;", expect: ["9"] },
    div_t_quotient => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "div_t r = div(10, 3); printf(\"%d %d\\n\", r.quot, r.rem); return 0;", expect: ["3 1"] },
    ldiv_t_quotient => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "ldiv_t r = ldiv(10L, 3L); printf(\"%ld\\n\", r.quot); return 0;", expect: ["3"] },
    lldiv_t_quotient => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "lldiv_t r = lldiv(10LL, 3LL); printf(\"%lld\\n\", r.rem); return 0;", expect: ["1"] },
    malloc_free => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "int *p = malloc(sizeof(int)); *p = 4; printf(\"%d\\n\", *p); free(p); return 0;", expect: ["4"] },
    calloc_zeroes => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "int *p = calloc(2, sizeof(int)); printf(\"%d\\n\", p[1]); free(p); return 0;", expect: ["0"] },
    realloc_grow => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "int *p = malloc(sizeof(int)); *p=1; p=realloc(p,2*sizeof(int)); p[1]=2; printf(\"%d\\n\", p[0]+p[1]); free(p); return 0;", expect: ["3"] },
    getenv_path => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "printf(\"%d\\n\", getenv(\"PATH\") != 0); return 0;", expect: ["1"] },
    rand_after_srand => { includes: ["<stdio.h>", "<stdlib.h>"], decls: "", body: "srand(1); int a = rand(); srand(1); int b = rand(); printf(\"%d\\n\", a==b); return 0;", expect: ["1"] },
    qsort_sorts => {
        includes: ["<stdio.h>", "<stdlib.h>"],
        decls: "int cmp(const void *a, const void *b) { return *(int*)a - *(int*)b; }",
        body: "int a[]={3,1,2}; qsort(a,3,sizeof(int),cmp); printf(\"%d\\n\", a[1]); return 0;",
        expect: ["2"]
    },
    bsearch_finds => {
        includes: ["<stdio.h>", "<stdlib.h>"],
        decls: "int cmp(const void *a, const void *b) { return *(int*)a - *(int*)b; }",
        body: "int a[]={1,3,5}; int key=3; int *p=bsearch(&key,a,3,sizeof(int),cmp); printf(\"%d\\n\", p ? *p : 0); return 0;",
        expect: ["3"]
    },
}

c_compile_cases! {
    aligned_alloc_compile => { includes: ["<stdlib.h>"], decls: "", body: "void *p = aligned_alloc(16, 32); free(p); return 0;" },
    reallocarray_compile => { includes: ["<stdlib.h>"], decls: "", body: "void *p = reallocarray(0, 2, 4); free(p); return 0;" },
    system_compile => { includes: ["<stdlib.h>"], decls: "", body: "return system(\"true\");" },
    abort_compile => { includes: ["<stdlib.h>"], decls: "", body: "if (0) abort(); return 0;" },
    exit_compile => { includes: ["<stdlib.h>"], decls: "", body: "if (0) exit(0); return 0;" },
    at_quick_exit_compile => { includes: ["<stdlib.h>"], decls: "void h(void) {}", body: "at_quick_exit(h); return 0;" },
}
