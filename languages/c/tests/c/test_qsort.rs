use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { declarations: $decls:expr, body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<stdlib.h>", "<string.h>"], $decls, $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    qsort_integers_ascending => {
        declarations: "int cmp_int(const void *a, const void *b) { return *(int*)a - *(int*)b; }",
        body: r#"
int arr[] = {5, 2, 8, 1, 9};
qsort(arr, 5, sizeof(int), cmp_int);
printf("%d %d %d %d %d\n", arr[0], arr[1], arr[2], arr[3], arr[4]);
return 0;
"#,
        expect: ["1 2 5 8 9"]
    },
    qsort_integers_descending => {
        declarations: "int cmp_desc(const void *a, const void *b) { return *(int*)b - *(int*)a; }",
        body: r#"
int arr[] = {3, 1, 4, 1, 5};
qsort(arr, 5, sizeof(int), cmp_desc);
printf("%d %d %d %d %d\n", arr[0], arr[1], arr[2], arr[3], arr[4]);
return 0;
"#,
        expect: ["5 4 3 1 1"]
    },
    qsort_already_sorted => {
        declarations: "int cmp_int(const void *a, const void *b) { return *(int*)a - *(int*)b; }",
        body: r#"
int arr[] = {1, 2, 3};
qsort(arr, 3, sizeof(int), cmp_int);
printf("%d %d %d\n", arr[0], arr[1], arr[2]);
return 0;
"#,
        expect: ["1 2 3"]
    },
    bsearch_finds_element => {
        declarations: "int cmp_int(const void *a, const void *b) { return *(int*)a - *(int*)b; }",
        body: r#"
int arr[] = {1, 3, 5, 7, 9};
int key = 5;
int *p = (int*)bsearch(&key, arr, 5, sizeof(int), cmp_int);
printf("%d\n", *p);
return 0;
"#,
        expect: ["5"]
    },
    bsearch_not_found_returns_null => {
        declarations: "int cmp_int(const void *a, const void *b) { return *(int*)a - *(int*)b; }",
        body: r#"
int arr[] = {2, 4, 6};
int key = 5;
int *p = (int*)bsearch(&key, arr, 3, sizeof(int), cmp_int);
printf("%d\n", p == NULL ? 1 : 0);
return 0;
"#,
        expect: ["1"]
    }
}
