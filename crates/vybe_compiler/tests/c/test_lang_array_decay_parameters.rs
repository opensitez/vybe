//! Array decay in parameters, multidimensional forms, and sizeof caller vs callee.


c_run_cases! {
    sizeof_array_in_caller_counts_elements => { includes: ["<stdio.h>"], decls: "", body: "int a[6]={0}; printf(\"%zu\\n\", sizeof a / sizeof a[0]); return 0;", expect: ["6"] },
    sizeof_param_pointer_is_pointer_size => {
        includes: ["<stdio.h>"],
        decls: "int elem_size(int a[]){ return (int)(sizeof a / sizeof a[0]); }",
        body: "int a[6]={0}; printf(\"%d\\n\", elem_size(a)); return 0;",
        expect: ["2"]
    },
    sizeof_param_with_brackets_still_pointer => {
        includes: ["<stdio.h>"],
        decls: "int elem_size(int a[6]){ return (int)(sizeof a / sizeof a[0]); }",
        body: "int a[6]={0}; printf(\"%d\\n\", elem_size(a)); return 0;",
        expect: ["2"]
    },
    callee_reads_second_via_decayed_param => {
        includes: ["<stdio.h>"],
        decls: "int second(int a[]){ return a[1]; }",
        body: "int a[3]={10,20,30}; printf(\"%d\\n\", second(a)); return 0;",
        expect: ["20"]
    },
    callee_mutates_through_decayed_param => {
        includes: ["<stdio.h>"],
        decls: "void set0(int a[]){ a[0]=99; }",
        body: "int a[2]={1,2}; set0(a); printf(\"%d\\n\", a[0]); return 0;",
        expect: ["99"]
    },
    const_array_param_reads_without_copy => {
        includes: ["<stdio.h>"],
        decls: "int sum2(const int a[]){ return a[0]+a[1]; }",
        body: "int a[2]={3,4}; printf(\"%d\\n\", sum2(a)); return 0;",
        expect: ["7"]
    },
    char_array_param_as_pointer_string => {
        includes: ["<stdio.h>"],
        decls: "int len(char a[]){ int n=0; while(a[n]) n++; return n; }",
        body: "char s[]=\"four\"; printf(\"%d\\n\", len(s)); return 0;",
        expect: ["4"]
    },
    multidim_second_row_via_row_pointer_param => {
        includes: ["<stdio.h>"],
        decls: "int corner(int a[][3]){ return a[1][2]; }",
        body: "int m[2][3]={{1,2,3},{4,5,6}}; printf(\"%d\\n\", corner(m)); return 0;",
        expect: ["6"]
    },
    multidim_first_row_sum_with_known_cols => {
        includes: ["<stdio.h>"],
        decls: "int row0_sum(int a[][4]){ return a[0][0]+a[0][1]+a[0][2]+a[0][3]; }",
        body: "int m[2][4]={{1,2,3,4},{0,0,0,0}}; printf(\"%d\\n\", row0_sum(m)); return 0;",
        expect: ["10"]
    },
    pointer_to_array_param_reads_row => {
        includes: ["<stdio.h>"],
        decls: "int first_of_row(int (*row)[3]){ return (*row)[0]; }",
        body: "int m[2][3]={{7,8,9},{1,2,3}}; printf(\"%d\\n\", first_of_row(&m[1])); return 0;",
        expect: ["1"]
    },
    sizeof_multidim_caller_full_array => { includes: ["<stdio.h>"], decls: "", body: "int m[2][3]={{0}}; printf(\"%zu\\n\", sizeof m); return 0;", expect: ["24"] },
    sizeof_multidim_param_is_pointer => {
        includes: ["<stdio.h>"],
        decls: "int sz(int a[][3]){ return (int)sizeof a; }",
        body: "int m[2][3]={{0}}; printf(\"%d\\n\", sz(m)); return 0;",
        expect: ["8"]
    },
    sizeof_static_sized_param_inside_function => {
        includes: ["<stdio.h>"],
        decls: "int n(int a[static 5]){ return (int)(sizeof a); }",
        body: "int a[5]={0}; printf(\"%d\\n\", n(a)); return 0;",
        expect: ["8"]
    },
    pass_array_slice_via_pointer_offset => {
        includes: ["<stdio.h>"],
        decls: "int head(int *p){ return p[0]; }",
        body: "int a[4]={11,22,33,44}; printf(\"%d\\n\", head(a+2)); return 0;",
        expect: ["33"]
    },
    return_total_using_length_param => {
        includes: ["<stdio.h>"],
        decls: "int total(int a[], int n){ int s=0; for(int i=0;i<n;i++) s+=a[i]; return s; }",
        body: "int a[4]={1,2,3,4}; printf(\"%d\\n\", total(a,4)); return 0;",
        expect: ["10"]
    },
    double_array_param_reads_value => {
        includes: ["<stdio.h>"],
        decls: "double pick(double a[], int i){ return a[i]; }",
        body: "double d[2]={1.5,2.5}; printf(\"%.1f\\n\", pick(d,1)); return 0;",
        expect: ["2.5"]
    },
    struct_array_param_reads_field => {
        includes: ["<stdio.h>"],
        decls: "struct P{int v;}; int get(struct P a[], int i){ return a[i].v; }",
        body: "struct P a[2]={{5},{6}}; printf(\"%d\\n\", get(a,1)); return 0;",
        expect: ["6"]
    },
    incomplete_array_param_declaration_compiles => {
        includes: ["<stdio.h>"],
        decls: "void touch(int a[]); void touch(int a[]){ a[0]=1; }",
        body: "int a[1]={0}; touch(a); printf(\"%d\\n\", a[0]); return 0;",
        expect: ["1"]
    },
    zero_length_caller_array_address_still_passes => {
        includes: ["<stdio.h>"],
        decls: "int is_null(int *p){ return p==0; }",
        body: "int *p=0; printf(\"%d\\n\", is_null(p)); return 0;",
        expect: ["1"]
    },
    typedef_array_param_same_as_pointer => {
        includes: ["<stdio.h>"],
        decls: "typedef int row_t[3]; int mid(row_t a){ return a[1]; }",
        body: "int a[3]={1,9,3}; printf(\"%d\\n\", mid(a)); return 0;",
        expect: ["9"]
    },
    nested_call_preserves_array_identity => {
        includes: ["<stdio.h>"],
        decls: "int id(int a[]){ return a[0]; } int wrap(int a[]){ return id(a); }",
        body: "int a[1]={42}; printf(\"%d\\n\", wrap(a)); return 0;",
        expect: ["42"]
    },
    sizeof_char_array_in_caller_includes_nul => { includes: ["<stdio.h>"], decls: "", body: "char s[]=\"ab\"; printf(\"%zu\\n\", sizeof s); return 0;", expect: ["3"] },
    sizeof_char_param_is_pointer => {
        includes: ["<stdio.h>"],
        decls: "int psz(char a[]){ return (int)sizeof a; }",
        body: "char s[]=\"ab\"; printf(\"%d\\n\", psz(s)); return 0;",
        expect: ["8"]
    },
    multidim_decay_passes_row_pointer => {
        includes: ["<stdio.h>"],
        decls: "int row1_first(int (*a)[2]){ return a[1][0]; }",
        body: "int m[2][2]={{1,2},{3,4}}; printf(\"%d\\n\", row1_first(m)); return 0;",
        expect: ["3"]
    },
    sizeof_caller_vs_callee_side_by_side => {
        includes: ["<stdio.h>"],
        decls: "int callee_words(int a[]){ return (int)(sizeof a); }",
        body: "int a[4]={0}; printf(\"%zu %d\\n\", sizeof a, callee_words(a)); return 0;",
        expect: ["16 8"]
    },
    static_global_array_size_in_caller => { includes: ["<stdio.h>"], decls: "static int ga[5]={0};", body: "printf(\"%zu\\n\", sizeof ga); return 0;", expect: ["20"] },
    static_global_passed_param_pointer_size => {
        includes: ["<stdio.h>"],
        decls: "static int ga[5]={0}; int psz(int a[]){ return (int)sizeof a; }",
        body: "printf(\"%d\\n\", psz(ga)); return 0;",
        expect: ["8"]
    },
    array_param_pointer_arithmetic_in_callee => {
        includes: ["<stdio.h>"],
        decls: "int third(int a[]){ int *p=a; return *(p+2); }",
        body: "int a[4]={1,2,3,4}; printf(\"%d\\n\", third(a)); return 0;",
        expect: ["3"]
    },
    const_char_param_printf => {
        includes: ["<stdio.h>"],
        decls: "void show(const char a[]){ printf(\"%s\\n\", a); }",
        body: "show(\"ok\"); return 0;",
        expect: ["ok"]
    },
    two_d_param_write_visible_in_caller => {
        includes: ["<stdio.h>"],
        decls: "void set(int a[][2]){ a[1][1]=77; }",
        body: "int m[2][2]={{0}}; set(m); printf(\"%d\\n\", m[1][1]); return 0;",
        expect: ["77"]
    },
    sizeof_short_array_caller => { includes: ["<stdio.h>"], decls: "", body: "short a[8]={0}; printf(\"%zu\\n\", sizeof a); return 0;", expect: ["16"] },
    sizeof_short_param_pointer => {
        includes: ["<stdio.h>"],
        decls: "int psz(short a[]){ return (int)sizeof a; }",
        body: "short a[8]={0}; printf(\"%d\\n\", psz(a)); return 0;",
        expect: ["8"]
    },
    sizeof_long_long_array_caller => { includes: ["<stdio.h>"], decls: "", body: "long long a[2]={0}; printf(\"%zu\\n\", sizeof a); return 0;", expect: ["16"] },
    sizeof_long_long_param_pointer => {
        includes: ["<stdio.h>"],
        decls: "int psz(long long a[]){ return (int)sizeof a; }",
        body: "long long a[2]={0}; printf(\"%d\\n\", psz(a)); return 0;",
        expect: ["8"]
    },
    function_returns_pointer_into_array => {
        includes: ["<stdio.h>"],
        decls: "int* at(int a[], int i){ return &a[i]; }",
        body: "int a[3]={5,6,7}; printf(\"%d\\n\", *at(a,2)); return 0;",
        expect: ["7"]
    },
    array_param_with_enum_length => {
        includes: ["<stdio.h>"],
        decls: "enum{N=3}; int last(int a[]){ return a[N-1]; }",
        body: "int a[3]={1,2,8}; printf(\"%d\\n\", last(a)); return 0;",
        expect: ["8"]
    },
    sizeof_param_does_not_see_caller_bound => {
        includes: ["<stdio.h>"],
        decls: "int words(int a[100]){ return (int)(sizeof a); }",
        body: "int a[3]={0}; printf(\"%d\\n\", words(a)); return 0;",
        expect: ["8"]
    },
    multidim_three_by_two_corner => {
        includes: ["<stdio.h>"],
        decls: "int val(int a[][2]){ return a[2][1]; }",
        body: "int m[3][2]={{1,2},{3,4},{5,6}}; printf(\"%d\\n\", val(m)); return 0;",
        expect: ["6"]
    },
    decayed_string_literal_to_const_param => {
        includes: ["<stdio.h>"],
        decls: "char first(const char a[]){ return a[0]; }",
        body: "printf(\"%c\\n\", first(\"Z\")); return 0;",
        expect: ["Z"]
    },
    array_param_same_storage_after_assign => {
        includes: ["<stdio.h>"],
        decls: "void bump(int a[]){ a[0]++; }",
        body: "int a[1]={4}; bump(a); bump(a); printf(\"%d\\n\", a[0]); return 0;",
        expect: ["6"]
    },
    sizeof_float_array_caller => { includes: ["<stdio.h>"], decls: "", body: "float a[5]={0}; printf(\"%zu\\n\", sizeof a); return 0;", expect: ["20"] },
    sizeof_float_param_pointer => {
        includes: ["<stdio.h>"],
        decls: "int psz(float a[]){ return (int)sizeof a; }",
        body: "float a[5]={0}; printf(\"%d\\n\", psz(a)); return 0;",
        expect: ["8"]
    },
    nested_multidim_param_inner_value => {
        includes: ["<stdio.h>"],
        decls: "int deep(int a[][2][2]){ return a[0][1][0]; }",
        body: "int m[1][2][2]={{{1,2},{3,4}}}; printf(\"%d\\n\", deep(m)); return 0;",
        expect: ["3"]
    },
    caller_element_size_matches_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int a[1]={0}; printf(\"%zu\\n\", sizeof a[0]); return 0;",
        expect: ["4"]
    },
    param_receives_address_of_compound_literal => {
        includes: ["<stdio.h>"],
        decls: "int pick(int a[]){ return a[1]; }",
        body: "printf(\"%d\\n\", pick((int[]){9,8,7})); return 0;",
        expect: ["8"]
    },
    sizeof_caller_two_dim_array => { includes: ["<stdio.h>"], decls: "", body: "int m[3][2]={{0}}; printf(\"%zu\\n\", sizeof m); return 0;", expect: ["24"] },
    sizeof_two_dim_param_row_pointer => {
        includes: ["<stdio.h>"],
        decls: "int psz(int a[][2]){ return (int)sizeof a; }",
        body: "int m[3][2]={{0}}; printf(\"%d\\n\", psz(m)); return 0;",
        expect: ["8"]
    },
}
