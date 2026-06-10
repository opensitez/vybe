use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>", "<stdlib.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    memcpy_copies_bytes => {
        body: r#"
int src[3] = {1, 2, 3};
int dst[3];
memcpy(dst, src, sizeof(src));
printf("%d %d %d\n", dst[0], dst[1], dst[2]);
return 0;
"#,
        expect: ["1 2 3"]
    },
    memcpy_string_content => {
        body: r#"
char src[] = "hello";
char dst[6];
memcpy(dst, src, 6);
printf("%s\n", dst);
return 0;
"#,
        expect: ["hello"]
    },
    memmove_overlapping => {
        body: r#"
char buf[] = "abcde";
memmove(buf + 1, buf, 4);
printf("%c%c%c%c%c\n", buf[0], buf[1], buf[2], buf[3], buf[4]);
return 0;
"#,
        expect: ["aabcd"]
    },
    memset_zero_fill => {
        body: r#"
int arr[4] = {1, 2, 3, 4};
memset(arr, 0, sizeof(arr));
printf("%d %d %d %d\n", arr[0], arr[1], arr[2], arr[3]);
return 0;
"#,
        expect: ["0 0 0 0"]
    },
    memset_byte_fill => {
        body: r#"
char buf[4];
memset(buf, 'X', 3);
buf[3] = '\0';
printf("%s\n", buf);
return 0;
"#,
        expect: ["XXX"]
    },
    memcmp_equal_buffers => {
        body: r#"
char a[] = "abc";
char b[] = "abc";
printf("%d\n", memcmp(a, b, 3));
return 0;
"#,
        expect: ["0"]
    },
    memcmp_unequal_buffers => {
        body: r#"
char a[] = "abc";
char b[] = "abd";
printf("%d\n", memcmp(a, b, 3) < 0 ? -1 : 1);
return 0;
"#,
        expect: ["-1"]
    },
    realloc_grows_buffer => {
        body: r#"
int *p = (int*)malloc(2 * sizeof(int));
p[0] = 1; p[1] = 2;
p = (int*)realloc(p, 4 * sizeof(int));
p[2] = 3; p[3] = 4;
printf("%d %d %d %d\n", p[0], p[1], p[2], p[3]);
free(p);
return 0;
"#,
        expect: ["1 2 3 4"]
    },
    memchr_finds_byte => {
        body: r#"
char buf[] = "hello";
char *p = (char*)memchr(buf, 'l', 5);
printf("%d\n", (int)(p - buf));
return 0;
"#,
        expect: ["2"]
    }
}
