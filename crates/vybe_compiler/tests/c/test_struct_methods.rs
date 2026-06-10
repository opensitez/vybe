use super::helpers::*;

// Struct + function patterns simulating OOP-style interfaces
macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<math.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    struct_with_function_pointer_member => {
        declarations: r#"
typedef struct {
    int value;
    int (*double_fn)(int);
} Widget;
int double_val(int x) { return x * 2; }
"#,
        body: "Widget w = {5, double_val};\nprintf(\"%d\\n\", w.double_fn(w.value));\nreturn 0;",
        expect: ["10"]
    },
    vtable_style_dispatch => {
        declarations: r#"
typedef struct {
    const char *(*name)(void);
    int (*area)(int, int);
} ShapeOps;
const char *rect_name(void) { return "rect"; }
int rect_area(int w, int h) { return w * h; }
"#,
        body: "ShapeOps ops = {rect_name, rect_area};\nprintf(\"%s %d\\n\", ops.name(), ops.area(3,4));\nreturn 0;",
        expect: ["rect 12"]
    },
    constructor_style_function => {
        declarations: r#"
typedef struct { int x; int y; int z; } Vec3;
Vec3 vec3_new(int x, int y, int z) { Vec3 v = {x, y, z}; return v; }
int vec3_dot(Vec3 a, Vec3 b) { return a.x*b.x + a.y*b.y + a.z*b.z; }
"#,
        body: "Vec3 a = vec3_new(1,2,3);\nVec3 b = vec3_new(4,5,6);\nprintf(\"%d\\n\", vec3_dot(a,b));\nreturn 0;",
        expect: ["32"]
    },
    struct_method_modifies_by_pointer => {
        declarations: r#"
typedef struct { int count; } Counter;
void counter_increment(Counter *c) { c->count++; }
void counter_reset(Counter *c) { c->count = 0; }
"#,
        body: "Counter c = {0};\ncounter_increment(&c);\ncounter_increment(&c);\ncounter_increment(&c);\nprintf(\"%d\\n\", c.count);\ncounter_reset(&c);\nprintf(\"%d\\n\", c.count);\nreturn 0;",
        expect: ["3", "0"]
    }
}
