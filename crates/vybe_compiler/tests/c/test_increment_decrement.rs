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
    prefix_increment_updates_before_use => { declarations: "int x = 4;", body: "printf(\"%d\\n\", ++x);\nreturn 0;", expect: ["5"] },
    postfix_increment_uses_old_value => { declarations: "int x = 4;", body: "printf(\"%d\\n\", x++);\nreturn 0;", expect: ["4"] },
    postfix_increment_updates_after_expression => { declarations: "int x = 4;", body: "x++;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["5"] },
    prefix_decrement_updates_before_use => { declarations: "int x = 4;", body: "printf(\"%d\\n\", --x);\nreturn 0;", expect: ["3"] },
    postfix_decrement_uses_old_value => { declarations: "int x = 4;", body: "printf(\"%d\\n\", x--);\nreturn 0;", expect: ["4"] },
    postfix_decrement_updates_after_expression => { declarations: "int x = 4;", body: "x--;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["3"] },
    prefix_increment_in_addition_uses_new_value => { declarations: "int x = 4;", body: "printf(\"%d\\n\", ++x + 1);\nreturn 0;", expect: ["6"] },
    postfix_increment_in_addition_uses_old_value => { declarations: "int x = 4;", body: "printf(\"%d\\n\", x++ + 1);\nreturn 0;", expect: ["5"] },
    prefix_decrement_in_addition_uses_new_value => { declarations: "int x = 4;", body: "printf(\"%d\\n\", --x + 1);\nreturn 0;", expect: ["4"] },
    postfix_decrement_in_addition_uses_old_value => { declarations: "int x = 4;", body: "printf(\"%d\\n\", x-- + 1);\nreturn 0;", expect: ["5"] },
    increment_can_drive_loop_counter_manual => { declarations: "int x = 0;", body: "printf(\"%d\\n\", x++);\nprintf(\"%d\\n\", x++);\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["0", "1", "2"] },
    decrement_can_walk_value_down => { declarations: "int x = 3;", body: "printf(\"%d\\n\", x--);\nprintf(\"%d\\n\", --x);\nreturn 0;", expect: ["3", "1"] },
    increment_on_char_advances_ascii_code => { declarations: "char c = 'a';", body: "++c;\nprintf(\"%c\\n\", c);\nreturn 0;", expect: ["b"] },
    decrement_on_char_rewinds_ascii_code => { declarations: "char c = 'b';", body: "--c;\nprintf(\"%c\\n\", c);\nreturn 0;", expect: ["a"] },
    prefix_increment_value_can_be_assigned => { declarations: "int x = 4; int y = 0;", body: "y = ++x;\nprintf(\"%d %d\\n\", x, y);\nreturn 0;", expect: ["5 5"] },
    postfix_increment_value_can_be_assigned => { declarations: "int x = 4; int y = 0;", body: "y = x++;\nprintf(\"%d %d\\n\", x, y);\nreturn 0;", expect: ["5 4"] },
    prefix_decrement_value_can_be_assigned => { declarations: "int x = 4; int y = 0;", body: "y = --x;\nprintf(\"%d %d\\n\", x, y);\nreturn 0;", expect: ["3 3"] },
    postfix_decrement_value_can_be_assigned => { declarations: "int x = 4; int y = 0;", body: "y = x--;\nprintf(\"%d %d\\n\", x, y);\nreturn 0;", expect: ["3 4"] },
    increment_in_condition_uses_new_value => { declarations: "int x = 0;", body: "if (++x) puts(\"true\"); else puts(\"false\");\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["true", "1"] },
    decrement_in_condition_can_make_zero_false => { declarations: "int x = 1;", body: "if (--x) puts(\"true\"); else puts(\"false\");\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["false", "0"] },
    increment_on_array_element_updates_slot => { declarations: "int arr[2] = {1, 2};", body: "arr[0]++;\nprintf(\"%d\\n\", arr[0]);\nreturn 0;", expect: ["2"] },
    decrement_on_array_element_updates_slot => { declarations: "int arr[2] = {1, 2};", body: "--arr[1];\nprintf(\"%d\\n\", arr[1]);\nreturn 0;", expect: ["1"] },
    increment_and_decrement_can_cancel => { declarations: "int x = 9;", body: "x++;\nx--;\nprintf(\"%d\\n\", x);\nreturn 0;", expect: ["9"] },
    postfix_increment_on_double_uses_old_value => { declarations: "double x = 1.5;", body: "printf(\"%.1f\\n\", x++);\nprintf(\"%.1f\\n\", x);\nreturn 0;", expect: ["1.5", "2.5"] },
    prefix_increment_on_double_uses_new_value => { declarations: "double x = 1.5;", body: "printf(\"%.1f\\n\", ++x);\nreturn 0;", expect: ["2.5"] },
    postfix_decrement_on_double_uses_old_value => { declarations: "double x = 2.5;", body: "printf(\"%.1f\\n\", x--);\nprintf(\"%.1f\\n\", x);\nreturn 0;", expect: ["2.5", "1.5"] }
}