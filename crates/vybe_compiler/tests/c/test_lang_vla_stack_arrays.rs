//! VLA stack arrays — sizeof, initialization, function passing (distinct from test_vla.rs).


c_run_cases! {
    vla_sizeof_matches_element_count_times_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=5; int a[n]; printf(\"%zu\\n\", sizeof a); return 0;",
        expect: ["20"]
    },
    vla_sizeof_with_char_elements => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=8; char a[n]; printf(\"%zu\\n\", sizeof a); return 0;",
        expect: ["8"]
    },
    vla_sizeof_doubles_runtime_length => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; double a[n]; printf(\"%zu\\n\", sizeof a); return 0;",
        expect: ["24"]
    },
    vla_element_assignment_and_read => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=4; int a[n]; a[3]=77; printf(\"%d\\n\", a[3]); return 0;",
        expect: ["77"]
    },
    vla_zero_initializes_via_braces => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; int a[n]={0}; printf(\"%d\\n\", a[2]); return 0;",
        expect: ["0"]
    },
    vla_partial_brace_initializer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=4; int a[n]={1,2}; printf(\"%d %d\\n\", a[0], a[3]); return 0;",
        expect: ["1 0"]
    },
    vla_passed_to_function_by_pointer_decay => {
        includes: ["<stdio.h>"],
        decls: "int sum_n(int *p, int n){ int s=0; for(int i=0;i<n;i++) s+=p[i]; return s; }",
        body: "int n=3; int a[n]; a[0]=1;a[1]=2;a[2]=3; printf(\"%d\\n\", sum_n(a,n)); return 0;",
        expect: ["6"]
    },
    vla_function_parameter_c99_form => {
        includes: ["<stdio.h>"],
        decls: "int last_elem(int n, int a[n]){ return a[n-1]; }",
        body: "int vals[4]={10,20,30,40}; printf(\"%d\\n\", last_elem(4, vals)); return 0;",
        expect: ["40"]
    },
    vla_multidim_row_access => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int r=2,c=3; int m[r][c]; m[1][2]=11; printf(\"%d\\n\", m[1][2]); return 0;",
        expect: ["11"]
    },
    vla_multidim_sizeof_total => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int r=2,c=3; int m[r][c]; printf(\"%zu\\n\", sizeof m); return 0;",
        expect: ["24"]
    },
    vla_length_from_variable_expression => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int base=2; int n=base+3; int a[n]; a[4]=9; printf(\"%d\\n\", a[4]); return 0;",
        expect: ["9"]
    },
    vla_declared_after_other_locals => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int x=1; int n=3; int a[n]; a[0]=x+4; printf(\"%d\\n\", a[0]); return 0;",
        expect: ["5"]
    },
    vla_in_for_loop_repeated_allocation => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int total=0; for(int k=1;k<=3;k++){ int a[k]; for(int i=0;i<k;i++) a[i]=i+1; total+=a[k-1]; } printf(\"%d\\n\", total); return 0;",
        expect: ["6"]
    },
    vla_pointer_to_first_element => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; int a[n]; a[0]=5;a[1]=6;a[2]=7; int *p=a; printf(\"%d\\n\", p[1]); return 0;",
        expect: ["6"]
    },
    vla_copy_into_fixed_buffer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; int src[n]; src[0]=1;src[1]=2;src[2]=3; int dst[3]; for(int i=0;i<n;i++) dst[i]=src[i]; printf(\"%d\\n\", dst[2]); return 0;",
        expect: ["3"]
    },
    vla_reverse_in_place => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; int a[n]; a[0]=1;a[1]=2;a[2]=3; for(int i=0;i<n/2;i++){ int t=a[i]; a[i]=a[n-1-i]; a[n-1-i]=t; } printf(\"%d %d %d\\n\", a[0],a[1],a[2]); return 0;",
        expect: ["3 2 1"]
    },
    vla_max_element_scan => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=4; int a[n]; a[0]=3;a[1]=9;a[2]=1;a[3]=5; int m=a[0]; for(int i=1;i<n;i++) if(a[i]>m) m=a[i]; printf(\"%d\\n\", m); return 0;",
        expect: ["9"]
    },
    vla_nested_blocks_different_lengths => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int outer=0; { int n=2; int a[n]; a[1]=4; outer=a[1]; } { int m=3; int b[m]; b[2]=1; outer+=b[2]; } printf(\"%d\\n\", outer); return 0;",
        expect: ["5"]
    },
    vla_size_from_function_return => {
        includes: ["<stdio.h>"],
        decls: "int width(void){ return 4; }",
        body: "int n=width(); int a[n]; a[3]=12; printf(\"%d\\n\", a[3]); return 0;",
        expect: ["12"]
    },
    vla_fill_with_index_pattern => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=5; int a[n]; for(int i=0;i<n;i++) a[i]=i*i; printf(\"%d\\n\", a[4]); return 0;",
        expect: ["16"]
    },
    vla_pass_row_to_helper => {
        includes: ["<stdio.h>"],
        decls: "int row_sum(int *row, int c){ int s=0; for(int i=0;i<c;i++) s+=row[i]; return s; }",
        body: "int r=2,c=2; int m[r][c]; m[0][0]=1;m[0][1]=2;m[1][0]=3;m[1][1]=4; printf(\"%d\\n\", row_sum(m[1],c)); return 0;",
        expect: ["7"]
    },
    vla_char_buffer_string_build => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=4; char s[n]; s[0]='a';s[1]='b';s[2]='c';s[3]='\\0'; printf(\"%s\\n\", s); return 0;",
        expect: ["abc"]
    },
    vla_bool_array_flags => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; _Bool f[n]; f[0]=1;f[1]=0;f[2]=1; int c=0; for(int i=0;i<n;i++) if(f[i]) c++; printf(\"%d\\n\", c); return 0;",
        expect: ["2"]
    },
    vla_unsigned_elements => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=2; unsigned a[n]; a[0]=5;a[1]=6; printf(\"%u\\n\", a[0]+a[1]); return 0;",
        expect: ["11"]
    },
    vla_short_elements_promoted_in_sum => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; short a[n]; a[0]=100;a[1]=200;a[2]=50; int s=a[0]+a[1]+a[2]; printf(\"%d\\n\", s); return 0;",
        expect: ["350"]
    },
    vla_length_one_single_element => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=1; int a[n]; a[0]=42; printf(\"%d\\n\", a[0]); return 0;",
        expect: ["42"]
    },
    vla_length_six_middling_size => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=6; int a[n]; for(int i=0;i<n;i++) a[i]=i; printf(\"%d\\n\", a[5]); return 0;",
        expect: ["5"]
    },
    vla_address_difference_in_elements => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=4; int a[n]; printf(\"%td\\n\", &a[2]-&a[0]); return 0;",
        expect: ["2"]
    },
    vla_subarray_via_pointer_offset => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=5; int a[n]; for(int i=0;i<n;i++) a[i]=i+1; int *mid=&a[2]; printf(\"%d\\n\", mid[1]); return 0;",
        expect: ["4"]
    },
    vla_write_then_read_prior_elements => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; int a[n]; a[2]=9; a[0]=1; a[1]=2; printf(\"%d %d\\n\", a[0], a[2]); return 0;",
        expect: ["1 9"]
    },
    vla_multidim_diagonal_init => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int r=3,c=3; int m[r][c]; for(int i=0;i<r;i++) for(int j=0;j<c;j++) m[i][j]=(i==j)?1:0; printf(\"%d\\n\", m[2][2]); return 0;",
        expect: ["1"]
    },
    vla_function_modifies_caller_buffer => {
        includes: ["<stdio.h>"],
        decls: "void dbl(int n, int a[n]){ for(int i=0;i<n;i++) a[i]*=2; }",
        body: "int n=2; int a[n]; a[0]=3;a[1]=4; dbl(n,a); printf(\"%d %d\\n\", a[0], a[1]); return 0;",
        expect: ["6 8"]
    },
    vla_sizeof_not_pointer_type => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=4; int a[n]; int *p=a; printf(\"%zu %zu\\n\", sizeof a, sizeof p); return 0;",
        expect: ["16 8"]
    },
    vla_conditional_length_via_ternary => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int flag=1; int n=flag?4:2; int a[n]; a[3]=8; printf(\"%d\\n\", a[3]); return 0;",
        expect: ["8"]
    },
    vla_aggregate_running_total => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=5; int a[n]; int s=0; for(int i=0;i<n;i++){ a[i]=i+1; s+=a[i]; } printf(\"%d\\n\", s); return 0;",
        expect: ["15"]
    },
    vla_swap_rows_in_matrix => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int r=2,c=2; int m[r][c]; m[0][0]=1;m[0][1]=2;m[1][0]=3;m[1][1]=4; for(int j=0;j<c;j++){ int t=m[0][j]; m[0][j]=m[1][j]; m[1][j]=t; } printf(\"%d %d\\n\", m[0][0], m[1][0]); return 0;",
        expect: ["3 1"]
    },
    vla_long_length_still_stack => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=10; int a[n]; a[9]=99; printf(\"%d\\n\", a[9]); return 0;",
        expect: ["99"]
    },
    vla_index_from_end_negative_offset => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=4; int a[n]; a[0]=10;a[1]=20;a[2]=30;a[3]=40; int *end=&a[n]; printf(\"%d\\n\", end[-1]); return 0;",
        expect: ["40"]
    },
    vla_compare_adjacent_elements => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=3; int a[n]; a[0]=1;a[1]=3;a[2]=2; int c=0; for(int i=0;i<n-1;i++) if(a[i]<a[i+1]) c++; printf(\"%d\\n\", c); return 0;",
        expect: ["1"]
    },
    vla_multidim_column_sum => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int r=2,c=3; int m[r][c]; for(int i=0;i<r;i++) for(int j=0;j<c;j++) m[i][j]=i+j; int col=0; for(int i=0;i<r;i++) col+=m[i][1]; printf(\"%d\\n\", col); return 0;",
        expect: ["3"]
    },
    vla_reuse_name_inner_shadow => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=2; int a[n]; a[0]=1; { int n=3; int a[n]; a[2]=7; printf(\"%d\\n\", a[2]); } printf(\"%d\\n\", a[0]); return 0;",
        expect: ["7", "1"]
    },
    vla_const_size_expression => {
        includes: ["<stdio.h>"],
        decls: "enum { MAX=5 };",
        body: "int n=MAX; int a[n]; a[4]=18; printf(\"%d\\n\", a[4]); return 0;",
        expect: ["18"]
    },
    vla_float_elements => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=2; float a[n]; a[0]=1.5f;a[1]=2.5f; printf(\"%.1f\\n\", a[0]+a[1]); return 0;",
        expect: ["4.0"]
    },
    vla_pass_to_variadic_style_manual => {
        includes: ["<stdio.h>"],
        decls: "void print3(int n, int *a){ printf(\"%d\\n\", a[0]+a[1]+a[2]); }",
        body: "int n=3; int a[n]; a[0]=2;a[1]=3;a[2]=4; print3(n,a); return 0;",
        expect: ["9"]
    },
    vla_binary_search_on_sorted => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n=5; int a[n]; a[0]=1;a[1]=3;a[2]=5;a[3]=7;a[4]=9; int t=5,lo=0,hi=n-1,mid; while(lo<=hi){ mid=(lo+hi)/2; if(a[mid]==t) break; else if(a[mid]<t) lo=mid+1; else hi=mid-1; } printf(\"%d\\n\", a[mid]); return 0;",
        expect: ["5"]
    },
}

c_compile_cases! {
    vla_parameter_prototype_compile => {
        includes: ["<stdio.h>"],
        decls: "void f(int n, int a[n]); void f(int n, int a[n]){}",
        body: "return 0;"
    },
    vla_multidim_parameter_compile => {
        includes: ["<stdio.h>"],
        decls: "void g(int r, int c, int m[r][c]); void g(int r, int c, int m[r][c]){}",
        body: "return 0;"
    },
}
