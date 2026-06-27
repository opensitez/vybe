//! sizeof and alignof expression semantics with numeric printed results.

use crate::helpers::*;

c_run_cases! {
    sizeof_short => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(short)); return 0;",
        expect: ["2"]
    },
    sizeof_long => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(long)); return 0;",
        expect: ["8"]
    },
    sizeof_long_long => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(long long)); return 0;",
        expect: ["8"]
    },
    sizeof_float => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(float)); return 0;",
        expect: ["4"]
    },
    sizeof_long_double => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(long double)); return 0;",
        expect: ["16"]
    },
    sizeof_bool => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(_Bool)); return 0;",
        expect: ["1"]
    },
    sizeof_signed_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(signed int)); return 0;",
        expect: ["4"]
    },
    sizeof_unsigned_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(unsigned int)); return 0;",
        expect: ["4"]
    },
    sizeof_int_array_ten => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(int[10])); return 0;",
        expect: ["40"]
    },
    sizeof_char_array_twenty => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(char[20])); return 0;",
        expect: ["20"]
    },
    sizeof_double_array_three => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(double[3])); return 0;",
        expect: ["24"]
    },
    sizeof_struct_member_int_field => {
        includes: ["<stdio.h>"],
        decls: "struct Box { int width; char tag; };",
        body: "struct Box b; printf(\"%d\\n\", (int)sizeof(b.width)); return 0;",
        expect: ["4"]
    },
    sizeof_struct_member_char_field => {
        includes: ["<stdio.h>"],
        decls: "struct Box { int width; char tag; };",
        body: "struct Box b; printf(\"%d\\n\", (int)sizeof(b.tag)); return 0;",
        expect: ["1"]
    },
    sizeof_struct_member_double_field => {
        includes: ["<stdio.h>"],
        decls: "struct Measure { double value; int id; };",
        body: "struct Measure m; printf(\"%d\\n\", (int)sizeof(m.value)); return 0;",
        expect: ["8"]
    },
    sizeof_struct_with_padding => {
        includes: ["<stdio.h>"],
        decls: "struct Pad { char a; int b; };",
        body: "printf(\"%d\\n\", (int)sizeof(struct Pad)); return 0;",
        expect: ["8"]
    },
    sizeof_nested_struct => {
        includes: ["<stdio.h>"],
        decls: "struct Inner { int x; }; struct Outer { struct Inner in; char c; };",
        body: "printf(\"%d\\n\", (int)sizeof(struct Outer)); return 0;",
        expect: ["8"]
    },
    sizeof_union_int_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; char c; double d; };",
        body: "union U u; printf(\"%d\\n\", (int)sizeof(u.i)); return 0;",
        expect: ["4"]
    },
    sizeof_union_double_member => {
        includes: ["<stdio.h>"],
        decls: "union U { int i; char c; double d; };",
        body: "union U u; printf(\"%d\\n\", (int)sizeof(u.d)); return 0;",
        expect: ["8"]
    },
    sizeof_enum_type => {
        includes: ["<stdio.h>"],
        decls: "enum Color { RED, GREEN, BLUE };",
        body: "printf(\"%d\\n\", (int)sizeof(enum Color)); return 0;",
        expect: ["4"]
    },
    sizeof_pointer_to_struct => {
        includes: ["<stdio.h>"],
        decls: "struct Node { int data; struct Node *next; };",
        body: "printf(\"%d\\n\", (int)sizeof(struct Node *)); return 0;",
        expect: ["8"]
    },
    sizeof_array_of_struct => {
        includes: ["<stdio.h>"],
        decls: "struct Pair { int a; int b; };",
        body: "printf(\"%d\\n\", (int)sizeof(struct Pair[4])); return 0;",
        expect: ["32"]
    },
    sizeof_static_int_array => {
        includes: ["<stdio.h>"],
        decls: "static int cache[7];",
        body: "printf(\"%d\\n\", (int)sizeof(cache)); return 0;",
        expect: ["28"]
    },
    sizeof_const_qualified_type => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(const double)); return 0;",
        expect: ["8"]
    },
    sizeof_volatile_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(volatile int)); return 0;",
        expect: ["4"]
    },
    sizeof_function_pointer_type => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(int (*)(double))); return 0;",
        expect: ["8"]
    },
    sizeof_does_not_evaluate_function_call => {
        includes: ["<stdio.h>"],
        decls: "int bump(int x) { return x + 1; }",
        body: "int n = 4; printf(\"%d %d\\n\", (int)sizeof(bump(n)), n); return 0;",
        expect: ["4 4"]
    },
    alignof_char => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(char)); return 0;",
        expect: ["1"]
    },
    alignof_short => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(short)); return 0;",
        expect: ["2"]
    },
    alignof_int => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(int)); return 0;",
        expect: ["4"]
    },
    alignof_long => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(long)); return 0;",
        expect: ["8"]
    },
    alignof_long_long => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(long long)); return 0;",
        expect: ["8"]
    },
    alignof_float => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(float)); return 0;",
        expect: ["4"]
    },
    alignof_double => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(double)); return 0;",
        expect: ["8"]
    },
    alignof_long_double => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(long double)); return 0;",
        expect: ["16"]
    },
    alignof_pointer => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(void *)); return 0;",
        expect: ["8"]
    },
    alignof_int_array => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(int[5])); return 0;",
        expect: ["4"]
    },
    alignof_struct_with_padding => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "struct Al { char a; int b; };",
        body: "printf(\"%d\\n\", (int)alignof(struct Al)); return 0;",
        expect: ["4"]
    },
    alignof_union => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "union Mix { int i; double d; };",
        body: "printf(\"%d\\n\", (int)alignof(union Mix)); return 0;",
        expect: ["8"]
    },
    alignof_enum => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "enum E { E0 = 0 };",
        body: "printf(\"%d\\n\", (int)alignof(enum E)); return 0;",
        expect: ["4"]
    },
    sizeof_and_alignof_int => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d %d\\n\", (int)sizeof(int), (int)alignof(int)); return 0;",
        expect: ["4 4"]
    },
    sizeof_struct_member_array_field => {
        includes: ["<stdio.h>"],
        decls: "struct Buf { char data[8]; };",
        body: "struct Buf b; printf(\"%d\\n\", (int)sizeof(b.data)); return 0;",
        expect: ["8"]
    },
    sizeof_zero_length_array_type => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(int[0])); return 0;",
        expect: ["0"]
    },
    sizeof_pointer_to_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(int *)); return 0;",
        expect: ["8"]
    },
    sizeof_pointer_to_char => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(char *)); return 0;",
        expect: ["8"]
    },
    sizeof_triple_pointer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(int **)); return 0;",
        expect: ["8"]
    },
    alignof_char_array => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(char[16])); return 0;",
        expect: ["1"]
    },
    sizeof_bitfield_struct_total => {
        includes: ["<stdio.h>"],
        decls: "struct Flags { unsigned x : 3; unsigned y : 5; };",
        body: "printf(\"%d\\n\", (int)sizeof(struct Flags)); return 0;",
        expect: ["4"]
    },
    sizeof_anonymous_struct_variable => {
        includes: ["<stdio.h>"],
        decls: "struct { int p; int q; } anon;",
        body: "printf(\"%d\\n\", (int)sizeof(anon)); return 0;",
        expect: ["8"]
    },
    sizeof_long_int => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)sizeof(long int)); return 0;",
        expect: ["8"]
    },
    alignof_function_pointer => {
        includes: ["<stdio.h>", "<stdalign.h>"],
        decls: "",
        body: "printf(\"%d\\n\", (int)alignof(int (*)(void))); return 0;",
        expect: ["8"]
    },
    sizeof_struct_two_chars => {
        includes: ["<stdio.h>"],
        decls: "struct Two { char a; char b; };",
        body: "printf(\"%d\\n\", (int)sizeof(struct Two)); return 0;",
        expect: ["2"]
    },
}

c_compile_cases! {
    sizeof_incomplete_struct_pointer => {
        includes: ["<stdio.h>"],
        decls: "struct Incomplete; struct Incomplete *p;",
        body: "return (int)sizeof(p);"
    },
    alignof_max_align_t => {
        includes: ["<stdalign.h>"],
        decls: "",
        body: "return (int)alignof(max_align_t);"
    },
}
