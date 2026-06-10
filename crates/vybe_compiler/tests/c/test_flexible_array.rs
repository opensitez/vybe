use super::helpers::*;

// C99 flexible array members (struct trailing [])
#[test]
fn flexible_array_member_declared() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
struct Buffer {
    int len;
    char data[];
};
int main() {
    struct Buffer *b = (struct Buffer*)malloc(sizeof(struct Buffer) + 6);
    b->len = 5;
    b->data[0] = 'h'; b->data[1] = 'e'; b->data[2] = 'l';
    b->data[3] = 'l'; b->data[4] = 'o'; b->data[5] = '\0';
    printf("%d %s\n", b->len, b->data);
    free(b);
    return 0;
}
"#,
        &["5 hello"],
    );
}

#[test]
fn flexible_array_struct_size_is_without_member() {
    assert_outputs(
        r#"
#include <stdio.h>
struct Flex { int n; double data[]; };
int main() {
    printf("%d\n", (int)sizeof(struct Flex));
    return 0;
}
"#,
        &["8"],
    );
}

#[test]
fn flexible_array_int_data() {
    assert_outputs(
        r#"
#include <stdio.h>
#include <stdlib.h>
struct IntVec { int n; int data[]; };
int main() {
    struct IntVec *v = (struct IntVec*)malloc(sizeof(struct IntVec) + 3 * sizeof(int));
    v->n = 3;
    v->data[0] = 10; v->data[1] = 20; v->data[2] = 30;
    for (int i = 0; i < v->n; i++) printf("%d\n", v->data[i]);
    free(v);
    return 0;
}
"#,
        &["10", "20", "30"],
    );
}
