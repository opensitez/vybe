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

#[test]
fn array_init_with_expression() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int n = 3;\nint arr[3] = {n, n+1, n+2};\nprintf(\"%d %d %d\\n\", arr[0], arr[1], arr[2]);\nreturn 0;",
        &["3 4 5"],
    );
}

#[test]
fn struct_init_aggregate() {
    assert_program(
        &["<stdio.h>"],
        "struct RGB { int r; int g; int b; };",
        "struct RGB color = {255, 128, 0};\nprintf(\"%d %d %d\\n\", color.r, color.g, color.b);\nreturn 0;",
        &["255 128 0"],
    );
}

#[test]
fn nested_struct_aggregate_init() {
    assert_program(
        &["<stdio.h>"],
        "struct Point { int x; int y; };\nstruct Line { struct Point a; struct Point b; };",
        "struct Line l = {{1,2},{3,4}};\nprintf(\"%d %d %d %d\\n\", l.a.x, l.a.y, l.b.x, l.b.y);\nreturn 0;",
        &["1 2 3 4"],
    );
}

c_cases! {
    static_array_with_string_elements => {
        declarations: "",
        body: r#"
static const char *MONTHS[] = {"Jan","Feb","Mar","Apr","May","Jun",
    "Jul","Aug","Sep","Oct","Nov","Dec"};
printf("%s %s\n", MONTHS[0], MONTHS[11]);
return 0;
"#,
        expect: ["Jan Dec"]
    },
    global_struct_array_init => {
        declarations: "struct Pt { int x; int y; };\nstruct Pt points[3] = {{1,2},{3,4},{5,6}};",
        body: "printf(\"%d %d\\n\", points[1].x, points[2].y);\nreturn 0;",
        expect: ["3 6"]
    },
    union_init_first_member => {
        declarations: "union Val { int i; float f; };",
        body: "union Val v = {42};\nprintf(\"%d\\n\", v.i);\nreturn 0;",
        expect: ["42"]
    }
}
