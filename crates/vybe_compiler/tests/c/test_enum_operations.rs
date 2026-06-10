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
    enum_comparison => {
        declarations: "enum Dir { NORTH=0, EAST=1, SOUTH=2, WEST=3 };",
        body: "enum Dir d = EAST;\nprintf(\"%d\\n\", d == EAST ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    enum_arithmetic => {
        declarations: "enum Num { ONE=1, TWO, THREE, FOUR };",
        body: "printf(\"%d\\n\", ONE + TWO + THREE + FOUR);\nreturn 0;",
        expect: ["10"]
    },
    enum_in_array_bounds => {
        declarations: "enum { SIZE = 5 };",
        body: "int arr[SIZE];\nfor (int i = 0; i < SIZE; i++) arr[i] = i;\nprintf(\"%d\\n\", arr[SIZE-1]);\nreturn 0;",
        expect: ["4"]
    },
    enum_to_string_lookup => {
        declarations: "enum Day { MON=0, TUE=1, WED=2 };\nconst char *day_names[] = {\"Monday\",\"Tuesday\",\"Wednesday\"};",
        body: "enum Day d = WED;\nprintf(\"%s\\n\", day_names[d]);\nreturn 0;",
        expect: ["Wednesday"]
    },
    enum_passed_to_function => {
        declarations: "enum Color { RED, GREEN, BLUE };\nvoid print_color(enum Color c) { printf(\"%d\\n\", c); }",
        body: "print_color(GREEN);\nreturn 0;",
        expect: ["1"]
    },
    enum_in_struct_field => {
        declarations: "enum State { OFF, ON };\nstruct Device { int id; enum State state; };",
        body: "struct Device dev = {1, ON};\nprintf(\"%d %d\\n\", dev.id, dev.state);\nreturn 0;",
        expect: ["1 1"]
    }
}
