use super::helpers::*;

// String-based I/O patterns: sprintf/sscanf/snprintf
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
    sprintf_builds_path => {
        body: r#"
char path[64];
sprintf(path, "/usr/%s/%d", "local", 42);
printf("%s\n", path);
return 0;
"#,
        expect: ["/usr/local/42"]
    },
    snprintf_safe_truncate => {
        body: r#"
char buf[8];
int n = snprintf(buf, sizeof(buf), "hello world");
printf("%s %d\n", buf, n);
return 0;
"#,
        expect: ["hello w 11"]
    },
    sscanf_parse_csv_line => {
        body: r#"
char name[16]; int age; float score;
sscanf("Alice,30,9.5", "%15[^,],%d,%f", name, &age, &score);
printf("%s %d %.1f\n", name, age, score);
return 0;
"#,
        expect: ["Alice 30 9.5"]
    },
    sscanf_parse_ip_address => {
        body: r#"
int a, b, c, d;
sscanf("192.168.1.1", "%d.%d.%d.%d", &a, &b, &c, &d);
printf("%d %d %d %d\n", a, b, c, d);
return 0;
"#,
        expect: ["192 168 1 1"]
    },
    sprintf_zero_pad => {
        body: r#"
char buf[12];
sprintf(buf, "%06d", 42);
printf("%s\n", buf);
return 0;
"#,
        expect: ["000042"]
    },
    sprintf_build_json_like => {
        body: r#"
char buf[64];
int id = 1;
const char *name = "item";
sprintf(buf, "{id:%d,name:%s}", id, name);
printf("%s\n", buf);
return 0;
"#,
        expect: ["{id:1,name:item}"]
    }
}
