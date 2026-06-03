use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    union_can_store_and_read_integer_member => { declarations: "union Data { int i; char c; };", body: "union Data data; data.i = 65; printf(\"%d\\n\", data.i);\nreturn 0;", expect: ["65"] },
    union_can_store_and_read_character_member => { declarations: "union Data { int i; char c; };", body: "union Data data; data.c = 'A'; printf(\"%c\\n\", data.c);\nreturn 0;", expect: ["A"] },
    union_size_matches_largest_member => { declarations: "union Data { int i; char c; };", body: "printf(\"%d\\n\", (int)sizeof(union Data));\nreturn 0;", expect: ["4"] },
    union_assignment_copies_stored_bits => { declarations: "union Data { int i; char c; };", body: "union Data first; first.i = 65; union Data second = first; printf(\"%d\\n\", second.i);\nreturn 0;", expect: ["65"] },
    union_pointer_can_read_integer_member => { declarations: "union Data { int i; char c; };", body: "union Data data; data.i = 77; union Data *p = &data; printf(\"%d\\n\", p->i);\nreturn 0;", expect: ["77"] },
    union_pointer_can_write_character_member => { declarations: "union Data { int i; char c; };", body: "union Data data; union Data *p = &data; p->c = 'B'; printf(\"%c\\n\", data.c);\nreturn 0;", expect: ["B"] },
    union_can_live_inside_struct_field => { declarations: "union Data { int i; char c; }; struct Box { union Data data; };", body: "struct Box box; box.data.i = 90; printf(\"%d\\n\", box.data.i);\nreturn 0;", expect: ["90"] },
    typedef_union_can_declare_variable => { declarations: "typedef union { int i; char c; } Data;", body: "Data data; data.i = 100; printf(\"%d\\n\", data.i);\nreturn 0;", expect: ["100"] },
    union_can_hold_double_member => { declarations: "union Value { double d; int i; };", body: "union Value value; value.d = 2.5; printf(\"%.1f\\n\", value.d);\nreturn 0;", expect: ["2.5"] },
    union_member_address_can_be_taken => { declarations: "union Data { int i; char c; };", body: "union Data data; data.i = 33; int *p = &data.i; printf(\"%d\\n\", *p);\nreturn 0;", expect: ["33"] },
    union_array_can_store_multiple_values => { declarations: "union Data { int i; char c; };", body: "union Data items[2]; items[0].i = 1; items[1].i = 2; printf(\"%d %d\\n\", items[0].i, items[1].i);\nreturn 0;", expect: ["1 2"] },
    union_can_be_passed_to_function_by_value => { declarations: "union Data { int i; char c; }; int read_i(union Data data) { return data.i; }", body: "union Data data; data.i = 55; printf(\"%d\\n\", read_i(data));\nreturn 0;", expect: ["55"] },
    union_can_be_returned_from_function => { declarations: "union Data { int i; char c; }; union Data make_data(int x) { union Data data; data.i = x; return data; }", body: "union Data data = make_data(66); printf(\"%d\\n\", data.i);\nreturn 0;", expect: ["66"] },
    union_field_can_be_overwritten_by_other_member => { declarations: "union Data { int i; char c; };", body: "union Data data; data.i = 65; data.c = 'C'; printf(\"%c\\n\", data.c);\nreturn 0;", expect: ["C"] },
    union_pointer_can_be_reassigned_between_values => { declarations: "union Data { int i; char c; };", body: "union Data first; union Data second; first.i = 1; second.i = 2; union Data *p = &first; p = &second; printf(\"%d\\n\", p->i);\nreturn 0;", expect: ["2"] },
    union_inside_array_can_be_indexed => { declarations: "union Data { int i; char c; };", body: "union Data items[3]; items[2].i = 9; printf(\"%d\\n\", items[2].i);\nreturn 0;", expect: ["9"] },
    union_member_can_participate_in_arithmetic => { declarations: "union Data { int i; char c; };", body: "union Data data; data.i = 7; printf(\"%d\\n\", data.i + 5);\nreturn 0;", expect: ["12"] },
    union_character_member_can_feed_putchar_style_format => { declarations: "union Data { int i; char c; };", body: "union Data data; data.c = 'D'; printf(\"%c\\n\", data.c);\nreturn 0;", expect: ["D"] },
    union_of_struct_and_int_can_read_struct_member => { declarations: "struct Pair { int a; int b; }; union Mixed { struct Pair pair; int i; };", body: "union Mixed mixed; mixed.pair.a = 3; mixed.pair.b = 4; printf(\"%d %d\\n\", mixed.pair.a, mixed.pair.b);\nreturn 0;", expect: ["3 4"] },
    union_member_size_can_be_read_via_sizeof => { declarations: "union Data { int i; char c; };", body: "union Data data; printf(\"%d %d\\n\", (int)sizeof(data.i), (int)sizeof(data.c));\nreturn 0;", expect: ["4 1"] },
    union_can_initialize_first_member_brace_style => { declarations: "union Data { int i; char c; };", body: "union Data data = {65}; printf(\"%d\\n\", data.i);\nreturn 0;", expect: ["65"] },
    union_can_be_nested_in_typedef_struct => { declarations: "typedef union { int i; char c; } Data; struct Holder { Data data; };", body: "struct Holder holder; holder.data.i = 44; printf(\"%d\\n\", holder.data.i);\nreturn 0;", expect: ["44"] },
    union_field_can_be_compound_assigned => { declarations: "union Data { int i; char c; };", body: "union Data data; data.i = 5; data.i += 2; printf(\"%d\\n\", data.i);\nreturn 0;", expect: ["7"] },
    union_pointer_to_member_address_can_update_storage => { declarations: "union Data { int i; char c; };", body: "union Data data; int *p = &data.i; *p = 88; printf(\"%d\\n\", data.i);\nreturn 0;", expect: ["88"] },
    union_with_pointer_member_can_store_string => { declarations: "union View { char *text; int i; };", body: "union View view; view.text = \"vybe\"; puts(view.text);\nreturn 0;", expect: ["vybe"] }
}