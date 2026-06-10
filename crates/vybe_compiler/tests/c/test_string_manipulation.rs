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
    string_reverse_manual => {
        body: r#"
char s[] = "hello";
int len = strlen(s);
for (int i = 0; i < len/2; i++) {
    char t = s[i]; s[i] = s[len-1-i]; s[len-1-i] = t;
}
printf("%s\n", s);
return 0;
"#,
        expect: ["olleh"]
    },
    string_count_occurrences => {
        body: r#"
char s[] = "banana";
int count = 0;
for (int i = 0; s[i]; i++) if (s[i] == 'a') count++;
printf("%d\n", count);
return 0;
"#,
        expect: ["3"]
    },
    string_remove_char => {
        body: r#"
char s[] = "hello world";
int w = 0;
for (int r = 0; s[r]; r++) if (s[r] != 'l') s[w++] = s[r];
s[w] = '\0';
printf("%s\n", s);
return 0;
"#,
        expect: ["heo word"]
    },
    sprintf_into_buffer => {
        body: r#"
char buf[32];
sprintf(buf, "val=%d", 42);
printf("%s\n", buf);
return 0;
"#,
        expect: ["val=42"]
    },
    snprintf_limits_length => {
        body: r#"
char buf[6];
snprintf(buf, sizeof(buf), "hello world");
printf("%s\n", buf);
return 0;
"#,
        expect: ["hello"]
    },
    string_is_palindrome => {
        body: r#"
char s[] = "racecar";
int len = strlen(s), ok = 1;
for (int i = 0; i < len/2; i++) if (s[i] != s[len-1-i]) { ok = 0; break; }
printf("%d\n", ok);
return 0;
"#,
        expect: ["1"]
    },
    string_to_uppercase_manual => {
        body: r#"
char s[] = "hello";
for (int i = 0; s[i]; i++) if (s[i] >= 'a' && s[i] <= 'z') s[i] -= 32;
printf("%s\n", s);
return 0;
"#,
        expect: ["HELLO"]
    },
    string_word_count => {
        body: r#"
char s[] = "one two three four";
int words = 0, in_word = 0;
for (int i = 0; s[i]; i++) {
    if (s[i] != ' ') { if (!in_word) { words++; in_word = 1; } }
    else in_word = 0;
}
printf("%d\n", words);
return 0;
"#,
        expect: ["4"]
    }
}
