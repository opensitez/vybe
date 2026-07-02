//! Nested structs, anonymous structs, and multi-level member access paths.


c_run_cases! {
    triple_nested_dot_path => {
        includes: ["<stdio.h>"],
        decls: "struct L3 { int v; }; struct L2 { struct L3 l3; }; struct L1 { struct L2 l2; };",
        body: "struct L1 o = {{{9}}}; printf(\"%d\\n\", o.l2.l3.v); return 0;",
        expect: ["9"]
    },
    triple_nested_arrow_path => {
        includes: ["<stdio.h>"],
        decls: "struct L3 { int v; }; struct L2 { struct L3 l3; }; struct L1 { struct L2 l2; };",
        body: "struct L1 o = {{{4}}}; struct L1 *p = &o; printf(\"%d\\n\", p->l2.l3.v); return 0;",
        expect: ["4"]
    },
    nested_struct_designated_init_dot => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int a; int b; }; struct Outer { struct Inner in; int tag; };",
        body: "struct Outer o = {.in = {.a = 2, .b = 3}, .tag = 1}; printf(\"%d %d %d\\n\", o.in.a, o.in.b, o.tag); return 0;",
        expect: ["2 3 1"]
    },
    nested_struct_designated_inner_only => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int x; int y; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {.in.y = 7}; printf(\"%d %d\\n\", o.in.x, o.in.y); return 0;",
        expect: ["0 7"]
    },
    anonymous_struct_in_struct_direct_access => {
        includes: ["<stdio.h>"],
        decls: "struct Pair { struct { int lo; int hi; } range; };",
        body: "struct Pair p; p.range.lo = 3; p.range.hi = 8; printf(\"%d %d\\n\", p.range.lo, p.range.hi); return 0;",
        expect: ["3 8"]
    },
    anonymous_struct_brace_init_in_outer => {
        includes: ["<stdio.h>"],
        decls: "struct Pair { struct { int lo; int hi; } range; };",
        body: "struct Pair p = {{5, 6}}; printf(\"%d %d\\n\", p.range.lo, p.range.hi); return 0;",
        expect: ["5 6"]
    },
    anonymous_struct_pointer_arrow => {
        includes: ["<stdio.h>"],
        decls: "struct Pair { struct { int lo; int hi; } range; };",
        body: "struct Pair p = {{1, 2}}; struct Pair *pp = &p; pp->range.hi = 11; printf(\"%d\\n\", p.range.hi); return 0;",
        expect: ["11"]
    },
    struct_contains_nested_struct_array => {
        includes: ["<stdio.h>"],
        decls: "struct Pt { int x; int y; }; struct Poly { struct Pt pts[2]; };",
        body: "struct Poly p = {{{1,2},{3,4}}}; printf(\"%d %d\\n\", p.pts[1].x, p.pts[0].y); return 0;",
        expect: ["3 2"]
    },
    nested_array_member_via_pointer => {
        includes: ["<stdio.h>"],
        decls: "struct Pt { int x; }; struct Row { struct Pt cells[3]; };",
        body: "struct Row r = {{{1},{2},{3}}}; struct Row *rp = &r; printf(\"%d\\n\", rp->cells[2].x); return 0;",
        expect: ["3"]
    },
    struct_in_struct_assignment_copies_inner => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer a = {{5}}; struct Outer b = a; b.in.n = 9; printf(\"%d %d\\n\", a.in.n, b.in.n); return 0;",
        expect: ["5 9"]
    },
    inner_struct_mutate_via_outer_pointer => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{1}}; struct Inner *ip = &o.in; ip->n = 6; printf(\"%d\\n\", o.in.n); return 0;",
        expect: ["6"]
    },
    nested_typedef_struct_access => {
        includes: ["<stdio.h>"],
        decls: "typedef struct { int v; } Inner; struct Outer { Inner in; };",
        body: "struct Outer o = {{12}}; printf(\"%d\\n\", o.in.v); return 0;",
        expect: ["12"]
    },
    doubly_nested_typedef_chain => {
        includes: ["<stdio.h>"],
        decls: "typedef struct { int v; } A; typedef struct { A a; } B; struct C { B b; };",
        body: "struct C c = {{{7}}}; printf(\"%d\\n\", c.b.a.v); return 0;",
        expect: ["7"]
    },
    function_returns_nested_struct_field_sum => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int a; int b; }; struct Outer { struct Inner in; }; int sum_outer(struct Outer o) { return o.in.a + o.in.b; }",
        body: "struct Outer o = {{2, 3}}; printf(\"%d\\n\", sum_outer(o)); return 0;",
        expect: ["5"]
    },
    function_takes_nested_struct_by_pointer => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; }; void bump(struct Outer *o) { o->in.n++; }",
        body: "struct Outer o = {{4}}; bump(&o); printf(\"%d\\n\", o.in.n); return 0;",
        expect: ["5"]
    },
    nested_struct_in_compound_literal => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer *p = &(struct Outer){.in = {.n = 15}}; printf(\"%d\\n\", p->in.n); return 0;",
        expect: ["15"]
    },
    four_level_nesting_access => {
        includes: ["<stdio.h>"],
        decls: "struct L4 { int v; }; struct L3 { struct L4 l4; }; struct L2 { struct L3 l3; }; struct L1 { struct L2 l2; };",
        body: "struct L1 o = {{{{42}}}}; printf(\"%d\\n\", o.l2.l3.l4.v); return 0;",
        expect: ["42"]
    },
    nested_struct_with_tag_and_anonymous_sibling => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { int tag; struct Inner body; };",
        body: "struct Outer o = {1, {2}}; printf(\"%d %d\\n\", o.tag, o.body.n); return 0;",
        expect: ["1 2"]
    },
    struct_nested_in_union_outer_wrapper => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; union U { struct Inner s; int i; }; struct Wrap { union U u; };",
        body: "struct Wrap w; w.u.s.n = 13; printf(\"%d\\n\", w.u.s.n); return 0;",
        expect: ["13"]
    },
    nested_member_address_taken_and_updated => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{0}}; int *p = &o.in.n; *p = 18; printf(\"%d\\n\", o.in.n); return 0;",
        expect: ["18"]
    },
    nested_struct_equality_via_fields => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int a; int b; }; struct Outer { struct Inner in; };",
        body: "struct Outer x = {{1,2}}; struct Outer y = {{1,2}}; printf(\"%d\\n\", x.in.a == y.in.a && x.in.b == y.in.b); return 0;",
        expect: ["1"]
    },
    nested_struct_in_global_storage => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; }; struct Outer g = {{21}};",
        body: "printf(\"%d\\n\", g.in.n); return 0;",
        expect: ["21"]
    },
    nested_struct_loop_mutate_inner_array => {
        includes: ["<stdio.h>"],
        decls: "struct Cell { int v; }; struct Grid { struct Cell row[3]; };",
        body: "struct Grid g = {{{1},{2},{3}}}; int i; for(i=0;i<3;i++) g.row[i].v *= 2; printf(\"%d %d\\n\", g.row[0].v, g.row[2].v); return 0;",
        expect: ["2 6"]
    },
    nested_struct_conditional_read => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int flag; int val; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{1, 99}}; printf(\"%d\\n\", o.in.flag ? o.in.val : 0); return 0;",
        expect: ["99"]
    },
    nested_anonymous_struct_two_groups => {
        includes: ["<stdio.h>"],
        decls: "struct Record { struct { int x; } a; struct { int y; } b; };",
        body: "struct Record r = {{3}, {4}}; printf(\"%d %d\\n\", r.a.x, r.b.y); return 0;",
        expect: ["3 4"]
    },
    nested_struct_copy_inner_only => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; int tag; };",
        body: "struct Outer a = {{5}, 1}; struct Inner copy = a.in; copy.n = 0; printf(\"%d %d\\n\", a.in.n, copy.n); return 0;",
        expect: ["5 0"]
    },
    nested_pointer_chain_double_arrow => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{6}}; struct Outer *op = &o; struct Inner *ip = &op->in; printf(\"%d\\n\", ip->n); return 0;",
        expect: ["6"]
    },
    nested_struct_ternary_select_field => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int a; int b; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{2, 7}}; printf(\"%d\\n\", 1 ? o.in.b : o.in.a); return 0;",
        expect: ["7"]
    },
    nested_struct_return_from_helper => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; }; struct Inner get_inner(struct Outer o) { return o.in; }",
        body: "struct Outer o = {{8}}; struct Inner i = get_inner(o); printf(\"%d\\n\", i.n); return 0;",
        expect: ["8"]
    },
    nested_struct_static_local_inner_read => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; }; int read_cached(void) { static struct Outer cache = {{10}}; return cache.in.n; }",
        body: "printf(\"%d\\n\", read_cached()); return 0;",
        expect: ["10"]
    },
    nested_struct_member_compound_add => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{3}}; o.in.n += 4; printf(\"%d\\n\", o.in.n); return 0;",
        expect: ["7"]
    },
    nested_struct_in_struct_array_of_outer => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer arr[2] = {{{1}}, {{2}}}; printf(\"%d %d\\n\", arr[0].in.n, arr[1].in.n); return 0;",
        expect: ["1 2"]
    },
    nested_struct_pointer_to_array_element_inner => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer arr[2] = {{{5}}, {{6}}}; struct Outer *p = &arr[1]; printf(\"%d\\n\", p->in.n); return 0;",
        expect: ["6"]
    },
    nested_anonymous_struct_designated_init => {
        includes: ["<stdio.h>"],
        decls: "struct Box { struct { int w; int h; } size; };",
        body: "struct Box b = {.size = {.w = 4, .h = 9}}; printf(\"%d %d\\n\", b.size.w, b.size.h); return 0;",
        expect: ["4 9"]
    },
    nested_struct_switch_on_inner_field => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int code; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{2}}; switch(o.in.code){case 2: printf(\"ok\\n\"); break; default: printf(\"no\\n\");} return 0;",
        expect: ["ok"]
    },
    nested_struct_bitwise_on_inner_field => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int mask; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{12}}; printf(\"%d\\n\", o.in.mask & 8); return 0;",
        expect: ["8"]
    },
    nested_struct_unary_minus_inner => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{5}}; printf(\"%d\\n\", -o.in.n); return 0;",
        expect: ["-5"]
    },
    nested_struct_postfix_increment_inner => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{1}}; int v = o.in.n++; printf(\"%d %d\\n\", v, o.in.n); return 0;",
        expect: ["1 2"]
    },
    nested_struct_prefix_increment_inner => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{1}}; int v = ++o.in.n; printf(\"%d %d\\n\", v, o.in.n); return 0;",
        expect: ["2 2"]
    },
    nested_struct_mixed_dot_and_paren => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int a; int b; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{3, 4}}; printf(\"%d\\n\", (o.in.a + o.in.b)); return 0;",
        expect: ["7"]
    },
    nested_struct_read_after_inner_reassign => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{1}}; o.in = (struct Inner){9}; printf(\"%d\\n\", o.in.n); return 0;",
        expect: ["9"]
    },
    nested_struct_two_inners_same_outer => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner left; struct Inner right; };",
        body: "struct Outer o = {{1, 2}}; printf(\"%d %d\\n\", o.left.n, o.right.n); return 0;",
        expect: ["1 2"]
    },
    nested_struct_pointer_to_left_inner => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner left; struct Inner right; };",
        body: "struct Outer o = {{3, 4}}; struct Inner *p = &o.left; p->n = 30; printf(\"%d %d\\n\", o.left.n, o.right.n); return 0;",
        expect: ["30 4"]
    },
    nested_struct_zero_init_inner_fields => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int a; int b; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {0}; printf(\"%d %d\\n\", o.in.a, o.in.b); return 0;",
        expect: ["0 0"]
    },
    nested_struct_char_and_int_inner => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { char c; int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{'Z', 26}}; printf(\"%c %d\\n\", o.in.c, o.in.n); return 0;",
        expect: ["Z 26"]
    },
    nested_struct_inner_array_sum_loop => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int vals[3]; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{{1,2,3}}}; int s=0,i; for(i=0;i<3;i++) s+=o.in.vals[i]; printf(\"%d\\n\", s); return 0;",
        expect: ["6"]
    },
    nested_struct_read_through_temp_copy => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; };",
        body: "struct Outer o = {{14}}; struct Inner tmp = o.in; printf(\"%d\\n\", tmp.n); return 0;",
        expect: ["14"]
    },
    nested_struct_outer_tag_preserved_when_inner_changes => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { int tag; struct Inner in; };",
        body: "struct Outer o = {7, {1}}; o.in.n = 2; printf(\"%d %d\\n\", o.tag, o.in.n); return 0;",
        expect: ["7 2"]
    },
    nested_anonymous_struct_in_typedef_outer => {
        includes: ["<stdio.h>"],
        decls: "typedef struct { struct { int x; int y; } pos; } Entity;",
        body: "Entity e = {{10, 20}}; printf(\"%d %d\\n\", e.pos.x, e.pos.y); return 0;",
        expect: ["10 20"]
    },
    nested_struct_function_pointer_reads_inner => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int n; }; struct Outer { struct Inner in; }; int pick(struct Outer o) { return o.in.n; }",
        body: "struct Outer o = {{17}}; printf(\"%d\\n\", pick(o)); return 0;",
        expect: ["17"]
    },
    nested_struct_deep_designated_partial => {
        includes: ["<stdio.h>"],
        decls: "struct L3 { int v; }; struct L2 { struct L3 l3; }; struct L1 { struct L2 l2; int k; };",
        body: "struct L1 o = {.l2.l3.v = 33, .k = 1}; printf(\"%d %d\\n\", o.l2.l3.v, o.k); return 0;",
        expect: ["33 1"]
    },
}
