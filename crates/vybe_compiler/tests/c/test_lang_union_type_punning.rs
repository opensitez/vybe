//! Union reads/writes, struct-wrapped unions, and observable integer results.


c_run_cases! {
    union_write_int_read_same_int => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 42; printf(\"%d\\n\", u.i); return 0;",
        expect: ["42"]
    },
    union_write_short_read_short => {
        includes: ["<stdio.h>"],
        decls: "union U { short s; };",
        body: "union U u; u.s = 17; printf(\"%d\\n\", (int)u.s); return 0;",
        expect: ["17"]
    },
    union_write_char_read_char => {
        includes: ["<stdio.h>"],
        decls: "union U { char c; };",
        body: "union U u; u.c = 'K'; printf(\"%d\\n\", (int)u.c); return 0;",
        expect: ["75"]
    },
    union_overwrite_int_with_new_int => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 1; u.i = 99; printf(\"%d\\n\", u.i); return 0;",
        expect: ["99"]
    },
    union_brace_init_first_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; char c; };",
        body: "union U u = {55}; printf(\"%d\\n\", u.i); return 0;",
        expect: ["55"]
    },
    union_assignment_copies_active_bits => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U a; union U b; a.i = 23; b = a; printf(\"%d\\n\", b.i); return 0;",
        expect: ["23"]
    },
    union_pointer_write_read_int => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; union U *p = &u; p->i = 31; printf(\"%d\\n\", u.i); return 0;",
        expect: ["31"]
    },
    union_member_address_update => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; int *p = &u.i; *p = 44; printf(\"%d\\n\", u.i); return 0;",
        expect: ["44"]
    },
    union_in_struct_write_int_field => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; char c; }; struct Box { union U data; };",
        body: "struct Box b; b.data.i = 60; printf(\"%d\\n\", b.data.i); return 0;",
        expect: ["60"]
    },
    union_in_struct_arrow_access => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; }; struct Box { union U data; };",
        body: "struct Box b; struct Box *bp = &b; bp->data.i = 61; printf(\"%d\\n\", b.data.i); return 0;",
        expect: ["61"]
    },
    named_union_member_in_struct => {
        includes: ["<stdio.h>"],
        decls: "union Payload { int i; short s; }; struct Msg { union Payload p; };",
        body: "struct Msg m; m.p.s = 7; printf(\"%d\\n\", (int)m.p.s); return 0;",
        expect: ["7"]
    },
    anonymous_union_in_struct_int_path => {
        includes: ["<stdio.h>"],
        decls: "struct Val { int tag; union { int i; short s; }; };",
        body: "struct Val v; v.i = 88; printf(\"%d\\n\", v.i); return 0;",
        expect: ["88"]
    },
    anonymous_union_in_struct_short_path => {
        includes: ["<stdio.h>"],
        decls: "struct Val { union { int i; short s; }; };",
        body: "struct Val v; v.s = 9; printf(\"%d\\n\", (int)v.s); return 0;",
        expect: ["9"]
    },
    struct_with_union_then_read_int_after_char_write => {
        includes: ["<stdio.h>"],
        decls: "struct S { union { int i; char c; } u; };",
        body: "struct S s; s.u.c = 'A'; printf(\"%d\\n\", (int)s.u.c); return 0;",
        expect: ["65"]
    },
    union_array_separate_elements => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U arr[3]; arr[0].i=1; arr[1].i=2; arr[2].i=3; printf(\"%d %d\\n\", arr[0].i, arr[2].i); return 0;",
        expect: ["1 3"]
    },
    union_passed_to_function_by_value => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; }; int read_u(union U u) { return u.i; }",
        body: "union U u; u.i = 70; printf(\"%d\\n\", read_u(u)); return 0;",
        expect: ["70"]
    },
    union_returned_from_function => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; }; union U make(int n) { union U u; u.i = n; return u; }",
        body: "union U u = make(71); printf(\"%d\\n\", u.i); return 0;",
        expect: ["71"]
    },
    union_compound_assignment_on_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 5; u.i += 3; printf(\"%d\\n\", u.i); return 0;",
        expect: ["8"]
    },
    union_nested_struct_member_write_read => {
        includes: ["<stdio.h>"],
        decls: "struct Pair { int a; int b; }; union U { struct Pair p; int i; };",
        body: "union U u; u.p.a = 2; u.p.b = 3; printf(\"%d %d\\n\", u.p.a, u.p.b); return 0;",
        expect: ["2 3"]
    },
    union_struct_member_via_pointer => {
        includes: ["<stdio.h>"],
        decls: "struct Pair { int a; }; union U { struct Pair p; };",
        body: "union U u; union U *pu = &u; pu->p.a = 11; printf(\"%d\\n\", u.p.a); return 0;",
        expect: ["11"]
    },
    union_typedef_alias_storage => {
        includes: ["<stdio.h>"],
        decls: "typedef union { int i; char c; } Data;",
        body: "Data d; d.i = 72; printf(\"%d\\n\", d.i); return 0;",
        expect: ["72"]
    },
    union_in_typedef_struct_holder => {
        includes: ["<stdio.h>"],
        decls: "typedef union { int i; } Data; struct Holder { Data d; };",
        body: "struct Holder h; h.d.i = 73; printf(\"%d\\n\", h.d.i); return 0;",
        expect: ["73"]
    },
    union_zero_init_then_assign => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u = {0}; u.i = 74; printf(\"%d\\n\", u.i); return 0;",
        expect: ["74"]
    },
    union_global_storage_read => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; }; union U g = {75};",
        body: "printf(\"%d\\n\", g.i); return 0;",
        expect: ["75"]
    },
    union_static_local_persists => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; }; int bump(void) { static union U u; u.i++; return u.i; }",
        body: "printf(\"%d\\n\", bump() + bump()); return 0;",
        expect: ["3"]
    },
    union_switch_on_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 2; switch(u.i){case 2: printf(\"hit\\n\"); break; default: printf(\"miss\\n\");} return 0;",
        expect: ["hit"]
    },
    union_equality_after_assignment => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U a, b; a.i = 5; b = a; printf(\"%d\\n\", a.i == b.i); return 0;",
        expect: ["1"]
    },
    union_member_in_arithmetic => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 6; printf(\"%d\\n\", u.i * 2); return 0;",
        expect: ["12"]
    },
    union_char_member_overwrites_then_read_char => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; char c; };",
        body: "union U u; u.i = 100; u.c = 'B'; printf(\"%d\\n\", (int)u.c); return 0;",
        expect: ["66"]
    },
    union_short_and_int_separate_members => {
        includes: ["<stdio.h>"],
        decls: "union U { short s; int i; };",
        body: "union U u; u.s = 3; printf(\"%d\\n\", (int)u.s); return 0;",
        expect: ["3"]
    },
    union_in_array_inside_struct => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; }; struct Bag { union U slots[2]; };",
        body: "struct Bag b; b.slots[0].i = 4; b.slots[1].i = 5; printf(\"%d %d\\n\", b.slots[0].i, b.slots[1].i); return 0;",
        expect: ["4 5"]
    },
    union_pointer_reassignment_between_objects => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U a, b; a.i = 1; b.i = 2; union U *p = &a; p = &b; printf(\"%d\\n\", p->i); return 0;",
        expect: ["2"]
    },
    union_postfix_increment_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 1; int v = u.i++; printf(\"%d %d\\n\", v, u.i); return 0;",
        expect: ["1 2"]
    },
    union_prefix_decrement_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 2; int v = --u.i; printf(\"%d %d\\n\", v, u.i); return 0;",
        expect: ["1 1"]
    },
    union_negation_of_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 8; printf(\"%d\\n\", -u.i); return 0;",
        expect: ["-8"]
    },
    union_bitwise_and_on_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 15; printf(\"%d\\n\", u.i & 8); return 0;",
        expect: ["8"]
    },
    union_ternary_on_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 0; printf(\"%d\\n\", u.i ? 1 : 2); return 0;",
        expect: ["2"]
    },
    union_struct_with_two_ints_sum => {
        includes: ["<stdio.h>"],
        decls: "struct Pair { int a; int b; }; union U { struct Pair p; int i; };",
        body: "union U u; u.p.a = 4; u.p.b = 5; printf(\"%d\\n\", u.p.a + u.p.b); return 0;",
        expect: ["9"]
    },
    union_nested_in_outer_struct_with_tag => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; }; struct Wrap { int tag; union U u; };",
        body: "struct Wrap w; w.tag = 1; w.u.i = 76; printf(\"%d %d\\n\", w.tag, w.u.i); return 0;",
        expect: ["1 76"]
    },
    union_read_after_struct_member_write_same_union => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; short s; }; struct S { union U u; };",
        body: "struct S s; s.u.i = 77; printf(\"%d\\n\", s.u.i); return 0;",
        expect: ["77"]
    },
    union_unsigned_member_store => {
        includes: ["<stdio.h>"],
        decls: "union U { unsigned u; };",
        body: "union U u; u.u = 78; printf(\"%u\\n\", u.u); return 0;",
        expect: ["78"]
    },
    union_long_member_store => {
        includes: ["<stdio.h>"],
        decls: "union U { long n; };",
        body: "union U u; u.n = 79; printf(\"%ld\\n\", u.n); return 0;",
        expect: ["79"]
    },
    union_in_conditional_expression => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 1; printf(\"%d\\n\", u.i > 0 ? 80 : 0); return 0;",
        expect: ["80"]
    },
    union_loop_accumulate_array => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U arr[3]; int s=0,i; for(i=0;i<3;i++){arr[i].i=i+1; s+=arr[i].i;} printf(\"%d\\n\", s); return 0;",
        expect: ["6"]
    },
    union_member_compare_less_than => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 3; printf(\"%d\\n\", u.i < 5); return 0;",
        expect: ["1"]
    },
    union_copy_then_mutate_source => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U a, b; a.i = 10; b = a; a.i = 20; printf(\"%d %d\\n\", a.i, b.i); return 0;",
        expect: ["20 10"]
    },
    union_struct_anonymous_in_union_outer => {
        includes: ["<stdio.h>"],
        decls: "union U { struct { int x; int y; } s; int i; };",
        body: "union U u; u.s.x = 6; u.s.y = 7; printf(\"%d %d\\n\", u.s.x, u.s.y); return 0;",
        expect: ["6 7"]
    },
    union_double_nested_struct_wrapper => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; }; struct Mid { union U u; }; struct Top { struct Mid m; };",
        body: "struct Top t; t.m.u.i = 81; printf(\"%d\\n\", t.m.u.i); return 0;",
        expect: ["81"]
    },
    union_read_through_const_pointer => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 82; const union U *p = &u; printf(\"%d\\n\", p->i); return 0;",
        expect: ["82"]
    },
    union_member_modulo_assign => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 10; u.i %= 3; printf(\"%d\\n\", u.i); return 0;",
        expect: ["1"]
    },
    union_xor_assign_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; };",
        body: "union U u; u.i = 12; u.i ^= 5; printf(\"%d\\n\", u.i); return 0;",
        expect: ["9"]
    },
}
