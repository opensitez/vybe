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
    bubble_sort_ints => {
        declarations: r#"
void bubble_sort(int *arr, int n) {
    for (int i = 0; i < n-1; i++)
        for (int j = 0; j < n-1-i; j++)
            if (arr[j] > arr[j+1]) {
                int t = arr[j]; arr[j] = arr[j+1]; arr[j+1] = t;
            }
}
"#,
        body: "int a[] = {5,3,8,1,9,2};\nbubble_sort(a, 6);\nprintf(\"%d %d %d %d %d %d\\n\", a[0],a[1],a[2],a[3],a[4],a[5]);\nreturn 0;",
        expect: ["1 2 3 5 8 9"]
    },
    insertion_sort_ints => {
        declarations: r#"
void insertion_sort(int *arr, int n) {
    for (int i = 1; i < n; i++) {
        int key = arr[i], j = i - 1;
        while (j >= 0 && arr[j] > key) { arr[j+1] = arr[j]; j--; }
        arr[j+1] = key;
    }
}
"#,
        body: "int a[] = {4,2,7,1,5};\ninsertion_sort(a, 5);\nprintf(\"%d %d %d %d %d\\n\", a[0],a[1],a[2],a[3],a[4]);\nreturn 0;",
        expect: ["1 2 4 5 7"]
    },
    binary_search_found => {
        declarations: r#"
int bsearch_idx(int *arr, int n, int target) {
    int lo=0, hi=n-1;
    while (lo <= hi) {
        int mid = (lo+hi)/2;
        if (arr[mid] == target) return mid;
        if (arr[mid] < target) lo = mid+1;
        else hi = mid-1;
    }
    return -1;
}
"#,
        body: "int a[] = {1,3,5,7,9,11};\nprintf(\"%d %d\\n\", bsearch_idx(a,6,7), bsearch_idx(a,6,4));\nreturn 0;",
        expect: ["3 -1"]
    },
    sieve_of_eratosthenes => {
        declarations: r#"
int is_prime[20];
void sieve(int n) {
    for (int i = 0; i < n; i++) is_prime[i] = 1;
    is_prime[0] = is_prime[1] = 0;
    for (int i = 2; i*i < n; i++)
        if (is_prime[i])
            for (int j = i*i; j < n; j += i) is_prime[j] = 0;
}
"#,
        body: "sieve(20);\nint count = 0;\nfor (int i = 0; i < 20; i++) if (is_prime[i]) count++;\nprintf(\"%d\\n\", count);\nreturn 0;",
        expect: ["8"]
    },
    max_subarray_kadane => {
        declarations: r#"
int max_subarray(int *a, int n) {
    int max_so_far = a[0], max_ending = a[0];
    for (int i = 1; i < n; i++) {
        max_ending = max_ending + a[i];
        if (max_ending < a[i]) max_ending = a[i];
        if (max_so_far < max_ending) max_so_far = max_ending;
    }
    return max_so_far;
}
"#,
        body: "int a[] = {-2, 1, -3, 4, -1, 2, 1, -5, 4};\nprintf(\"%d\\n\", max_subarray(a, 9));\nreturn 0;",
        expect: ["6"]
    },
    fibonacci_iterative => {
        declarations: r#"
int fib(int n) {
    if (n <= 1) return n;
    int a = 0, b = 1;
    for (int i = 2; i <= n; i++) { int c = a+b; a=b; b=c; }
    return b;
}
"#,
        body: "printf(\"%d %d %d\\n\", fib(10), fib(15), fib(20));\nreturn 0;",
        expect: ["55 610 6765"]
    }
}
