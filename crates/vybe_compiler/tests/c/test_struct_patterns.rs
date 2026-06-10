use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    struct_copy_is_deep => {
        declarations: "struct Val { int x; int y; };",
        body: "struct Val a = {1, 2};\nstruct Val b = a;\nb.x = 99;\nprintf(\"%d %d\\n\", a.x, b.x);\nreturn 0;",
        expect: ["1 99"]
    },
    struct_with_char_array => {
        declarations: "struct Name { char first[16]; char last[16]; };",
        body: "struct Name n;\nstrcpy(n.first, \"John\");\nstrcpy(n.last, \"Doe\");\nprintf(\"%s %s\\n\", n.first, n.last);\nreturn 0;",
        expect: ["John Doe"]
    },
    struct_in_array => {
        declarations: "typedef struct { int x; int y; } Point;",
        body: "Point pts[3] = {{0,0},{1,1},{2,4}};\nprintf(\"%d %d\\n\", pts[2].x, pts[2].y);\nreturn 0;",
        expect: ["2 4"]
    },
    struct_pointer_to_member_array => {
        declarations: "struct Vec { float v[3]; };",
        body: "struct Vec vv = {{1.0f, 2.0f, 3.0f}};\nfloat *p = vv.v;\nprintf(\"%.0f %.0f\\n\", p[0], p[2]);\nreturn 0;",
        expect: ["1 3"]
    },
    struct_linked_by_index => {
        declarations: "struct Item { int val; int next; };",
        body: r#"
struct Item list[3] = {{10, 1}, {20, 2}, {30, -1}};
int idx = 0;
while (idx >= 0) {
    printf("%d\n", list[idx].val);
    idx = list[idx].next;
}
return 0;
"#,
        expect: ["10", "20", "30"]
    },
    struct_equality_via_memcmp => {
        declarations: "struct Pair { int a; int b; };",
        body: "struct Pair p = {1, 2};\nstruct Pair q = {1, 2};\nprintf(\"%d\\n\", memcmp(&p, &q, sizeof(p)) == 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    }
}
