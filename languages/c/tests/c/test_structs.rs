use super::helpers::*;

#[test]
fn struct_basic() {
    let out = run_prints(
        r#"
#include <stdio.h>
struct Point {
    int x;
    int y;
};
int main() {
    struct Point p;
    p.x = 3;
    p.y = 4;
    printf("%d\n", p.x);
    printf("%d\n", p.y);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn struct_init() {
    let out = run_prints(
        r#"
#include <stdio.h>
struct Rect {
    int w;
    int h;
};
int main() {
    struct Rect r = {10, 5};
    printf("%d\n", r.w * r.h);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn typedef_struct() {
    let out = run_prints(
        r#"
#include <stdio.h>
typedef struct {
    int age;
    int score;
} Person;
int main() {
    Person p;
    p.age = 25;
    p.score = 100;
    printf("%d %d\n", p.age, p.score);
    return 0;
}
"#,
    );
    assert_eq!(out, vec!["25 100"]);
}
