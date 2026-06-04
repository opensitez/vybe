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
    nested_struct_member_access_reads_inner_field => { declarations: "struct Point { int x; int y; }; struct Box { struct Point origin; int size; };", body: "struct Box box = {{2, 3}, 4};\nprintf(\"%d %d %d\\n\", box.origin.x, box.origin.y, box.size);\nreturn 0;", expect: ["2 3 4"] },
    struct_pointer_arrow_reads_field => { declarations: "struct Point { int x; int y; };", body: "struct Point point = {3, 4}; struct Point *p = &point;\nprintf(\"%d\\n\", p->y);\nreturn 0;", expect: ["4"] },
    struct_pointer_arrow_writes_field => { declarations: "struct Point { int x; int y; };", body: "struct Point point = {3, 4}; struct Point *p = &point;\np->x = 9;\nprintf(\"%d\\n\", point.x);\nreturn 0;", expect: ["9"] },
    struct_assignment_copies_all_fields => { declarations: "struct Pair { int a; int b; };", body: "struct Pair first = {1, 2}; struct Pair second = first;\nprintf(\"%d %d\\n\", second.a, second.b);\nreturn 0;", expect: ["1 2"] },
    array_of_structs_can_be_indexed => { declarations: "struct Pair { int a; int b; };", body: "struct Pair pairs[2] = {{1, 2}, {3, 4}};\nprintf(\"%d %d\\n\", pairs[1].a, pairs[1].b);\nreturn 0;", expect: ["3 4"] },
    struct_can_be_passed_to_function_by_value => { declarations: "struct Pair { int a; int b; }; int sum_pair(struct Pair pair) { return pair.a + pair.b; }", body: "struct Pair pair = {5, 6};\nprintf(\"%d\\n\", sum_pair(pair));\nreturn 0;", expect: ["11"] },
    struct_can_be_returned_from_function => { declarations: "struct Pair { int a; int b; }; struct Pair make_pair(int a, int b) { struct Pair pair = {a, b}; return pair; }", body: "struct Pair pair = make_pair(7, 8);\nprintf(\"%d %d\\n\", pair.a, pair.b);\nreturn 0;", expect: ["7 8"] },
    struct_field_can_hold_array_pointer => { declarations: "struct Buffer { char *text; };", body: "struct Buffer buffer = {\"vybe\"};\nputs(buffer.text);\nreturn 0;", expect: ["vybe"] },
    nested_struct_pointer_arrow_can_follow_chain => { declarations: "struct Point { int x; int y; }; struct Box { struct Point origin; int size; };", body: "struct Box box = {{2, 3}, 4}; struct Box *p = &box;\nprintf(\"%d\\n\", p->origin.y);\nreturn 0;", expect: ["3"] },
    struct_field_assignment_can_use_other_field_expression => { declarations: "struct Pair { int a; int b; };", body: "struct Pair pair = {2, 0};\npair.b = pair.a + 5;\nprintf(\"%d\\n\", pair.b);\nreturn 0;", expect: ["7"] },
    sizeof_struct_reports_total_storage => { declarations: "struct Pair { int a; int b; };", body: "printf(\"%d\\n\", (int)sizeof(struct Pair));\nreturn 0;", expect: ["8"] },
    sizeof_struct_field_reports_member_storage => { declarations: "struct Pair { int a; int b; };", body: "struct Pair pair = {1, 2};\nprintf(\"%d\\n\", (int)sizeof(pair.a));\nreturn 0;", expect: ["4"] },
    struct_with_char_and_int_fields_can_initialize_both => { declarations: "struct Mixed { char c; int n; };", body: "struct Mixed value = {'A', 9};\nprintf(\"%c %d\\n\", value.c, value.n);\nreturn 0;", expect: ["A 9"] },
    struct_pointer_can_iterate_array_of_structs => { declarations: "struct Pair { int a; int b; };", body: "struct Pair pairs[2] = {{1, 2}, {3, 4}}; struct Pair *p = pairs;\nprintf(\"%d %d\\n\", p[0].b, p[1].a);\nreturn 0;", expect: ["2 3"] },
    struct_copy_is_independent_after_mutation => { declarations: "struct Pair { int a; int b; };", body: "struct Pair first = {1, 2}; struct Pair second = first; second.a = 9;\nprintf(\"%d %d\\n\", first.a, second.a);\nreturn 0;", expect: ["1 9"] },
    struct_member_address_can_be_taken_and_written => { declarations: "struct Pair { int a; int b; };", body: "struct Pair pair = {1, 2}; int *p = &pair.b; *p = 7;\nprintf(\"%d\\n\", pair.b);\nreturn 0;", expect: ["7"] },
    struct_initializer_can_nest_braces => { declarations: "struct Point { int x; int y; }; struct Box { struct Point origin; int size; };", body: "struct Box box = {{5, 6}, 7};\nprintf(\"%d %d %d\\n\", box.origin.x, box.origin.y, box.size);\nreturn 0;", expect: ["5 6 7"] },
    typedef_struct_value_can_be_declared_and_used => { declarations: "typedef struct { int x; int y; } Point;", body: "Point point = {8, 9};\nprintf(\"%d %d\\n\", point.x, point.y);\nreturn 0;", expect: ["8 9"] },
    struct_can_contain_array_member => { declarations: "struct Row { int values[3]; };", body: "struct Row row = {{2, 4, 6}};\nprintf(\"%d %d %d\\n\", row.values[0], row.values[1], row.values[2]);\nreturn 0;", expect: ["2 4 6"] },
    struct_pointer_to_array_member_can_index_values => { declarations: "struct Row { int values[3]; };", body: "struct Row row = {{2, 4, 6}}; struct Row *p = &row;\nprintf(\"%d\\n\", p->values[2]);\nreturn 0;", expect: ["6"] },
    struct_return_value_can_feed_function_argument => { declarations: "struct Pair { int a; int b; }; struct Pair make_pair(int a, int b) { struct Pair pair = {a, b}; return pair; } int sum_pair(struct Pair pair) { return pair.a + pair.b; }", body: "printf(\"%d\\n\", sum_pair(make_pair(3, 7)));\nreturn 0;", expect: ["10"] },
    struct_with_pointer_field_can_follow_suffix => { declarations: "struct Slice { char *text; };", body: "struct Slice slice = {\"hello\" + 2};\nputs(slice.text);\nreturn 0;", expect: ["llo"] },
    struct_array_element_can_be_updated_in_loop => { declarations: "struct Pair { int a; int b; };", body: "struct Pair pairs[2] = {{1, 2}, {3, 4}}; for (int i = 0; i < 2; i++) pairs[i].a += 10;\nprintf(\"%d %d\\n\", pairs[0].a, pairs[1].a);\nreturn 0;", expect: ["11 13"] },
    struct_pointer_can_be_reassigned_between_instances => { declarations: "struct Pair { int a; int b; };", body: "struct Pair first = {1, 2}; struct Pair second = {3, 4}; struct Pair *p = &first; p = &second;\nprintf(\"%d\\n\", p->b);\nreturn 0;", expect: ["4"] },
    struct_member_expression_can_drive_condition => { declarations: "struct Flag { int on; };", body: "struct Flag flag = {1}; if (flag.on) puts(\"on\"); else puts(\"off\");\nreturn 0;", expect: ["on"] },
    struct_field_can_store_function_pointer => { declarations: "int add_one(int x) { return x + 1; } struct Op { int (*apply)(int); };", body: "struct Op op = {add_one};\nprintf(\"%d\\n\", op.apply(8));\nreturn 0;", expect: ["9"] },
    nested_struct_copy_preserves_inner_fields => { declarations: "struct Point { int x; int y; }; struct Box { struct Point origin; int size; };", body: "struct Box first = {{1, 2}, 3}; struct Box second = first;\nprintf(\"%d %d %d\\n\", second.origin.x, second.origin.y, second.size);\nreturn 0;", expect: ["1 2 3"] },
    struct_member_can_be_compound_assigned => { declarations: "struct Pair { int a; int b; };", body: "struct Pair pair = {2, 3}; pair.b += 5;\nprintf(\"%d\\n\", pair.b);\nreturn 0;", expect: ["8"] },
    pointer_to_struct_array_can_advance_and_read_next => { declarations: "struct Pair { int a; int b; };", body: "struct Pair pairs[2] = {{1, 2}, {3, 4}}; struct Pair *p = pairs; p++;\nprintf(\"%d %d\\n\", p->a, p->b);\nreturn 0;", expect: ["3 4"] },
    struct_with_double_field_keeps_fractional_value => { declarations: "struct Measure { double value; };", body: "struct Measure measure = {2.5};\nprintf(\"%.1f\\n\", measure.value);\nreturn 0;", expect: ["2.5"] }
}
