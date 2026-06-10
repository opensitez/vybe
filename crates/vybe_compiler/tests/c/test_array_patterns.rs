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
fn array_size_inferred_from_initializer() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int arr[] = {1,2,3,4,5};\nprintf(\"%d\\n\", (int)(sizeof(arr)/sizeof(arr[0])));\nreturn 0;",
        &["5"],
    );
}

#[test]
fn array_partial_initialization_zeroes_rest() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int arr[5] = {1,2};\nprintf(\"%d %d %d\\n\", arr[0], arr[2], arr[4]);\nreturn 0;",
        &["1 0 0"],
    );
}

#[test]
fn array_pointer_subscript_equivalence() {
    assert_program(
        &["<stdio.h>"],
        "",
        "int arr[3] = {10,20,30};\nint *p = arr;\nprintf(\"%d %d\\n\", p[1], *(p+2));\nreturn 0;",
        &["20 30"],
    );
}

#[test]
fn string_array_of_pointers() {
    assert_program(
        &["<stdio.h>"],
        "",
        "const char *days[] = {\"Mon\",\"Tue\",\"Wed\"};\nprintf(\"%s %s\\n\", days[0], days[2]);\nreturn 0;",
        &["Mon Wed"],
    );
}

#[test]
fn array_reverse_in_place() {
    assert_program(
        &["<stdio.h>"],
        "",
        r#"
int arr[5] = {1,2,3,4,5};
int i=0, j=4;
while (i < j) { int t=arr[i]; arr[i]=arr[j]; arr[j]=t; i++; j--; }
printf("%d %d %d %d %d\n", arr[0], arr[1], arr[2], arr[3], arr[4]);
return 0;
"#,
        &["5 4 3 2 1"],
    );
}

#[test]
fn array_search_linear() {
    assert_program(
        &["<stdio.h>"],
        "",
        r#"
int arr[5] = {3,1,4,1,5};
int target = 4, found = -1;
for (int i = 0; i < 5; i++) if (arr[i] == target) { found = i; break; }
printf("%d\n", found);
return 0;
"#,
        &["2"],
    );
}

c_cases! {
    array_passed_as_pointer => {
        declarations: "int first(int *a) { return a[0]; }",
        body: "int arr[3] = {10,20,30};\nprintf(\"%d\\n\", first(arr));\nreturn 0;",
        expect: ["10"]
    },
    array_of_structs_sort_by_field => {
        declarations: "struct Item { int id; int val; };",
        body: r#"
struct Item items[3] = {{1,30},{2,10},{3,20}};
for (int i = 0; i < 2; i++)
    for (int j = 0; j < 2-i; j++)
        if (items[j].val > items[j+1].val) {
            struct Item t = items[j]; items[j]=items[j+1]; items[j+1]=t;
        }
printf("%d %d %d\n", items[0].val, items[1].val, items[2].val);
return 0;
"#,
        expect: ["10 20 30"]
    }
}
