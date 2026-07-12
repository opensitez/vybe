use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<string.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    strdup_copies_string => {
        body: r#"
char *s = strdup("hello");
printf("%s\n", s);
free(s);
return 0;
"#,
        expect: ["hello"]
    },
    strncpy_copies_n_chars => {
        body: r#"
char dst[10];
strncpy(dst, "hello world", 5);
dst[5] = '\0';
printf("%s\n", dst);
return 0;
"#,
        expect: ["hello"]
    },
    strncat_appends_n_chars => {
        body: r#"
char dst[20] = "hello";
strncat(dst, " world extra", 6);
printf("%s\n", dst);
return 0;
"#,
        expect: ["hello world"]
    },
    strncmp_compares_n_chars => {
        body: r#"
printf("%d\n", strncmp("abcdef", "abcxyz", 3));
printf("%d\n", strncmp("abcdef", "abcxyz", 4) < 0 ? -1 : 1);
return 0;
"#,
        expect: ["0", "-1"]
    },
    strpbrk_finds_first_char_in_set => {
        body: r#"
char *p = strpbrk("hello world", "aeiou");
printf("%s\n", p);
return 0;
"#,
        expect: ["ello world"]
    },
    strspn_counts_span => {
        body: r#"
size_t n = strspn("abc123", "abcdef");
printf("%d\n", (int)n);
return 0;
"#,
        expect: ["3"]
    },
    strcspn_counts_complement_span => {
        body: r#"
size_t n = strcspn("hello world", " \t");
printf("%d\n", (int)n);
return 0;
"#,
        expect: ["5"]
    },
    strstr_substring_found => {
        body: r#"
char *p = strstr("hello world", "world");
printf("%s\n", p);
return 0;
"#,
        expect: ["world"]
    },
    strstr_not_found_returns_null => {
        body: r#"
char *p = strstr("hello", "xyz");
printf("%d\n", p == NULL ? 1 : 0);
return 0;
"#,
        expect: ["1"]
    },
    strtok_basic_split => {
        body: r#"
char s[] = "a:b:c";
char *tok = strtok(s, ":");
while (tok) {
    printf("%s\n", tok);
    tok = strtok(NULL, ":");
}
return 0;
"#,
        expect: ["a", "b", "c"]
    }
}
