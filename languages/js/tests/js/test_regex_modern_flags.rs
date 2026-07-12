crate::js_cases! {
    regex_dotall_matches_newlines => {
        r#"
const re = /start.*end/s;
console.log(re.test("start\nend"));
"#,
        ["true"]
    };

    regex_dotall_reports_property => {
        r#"
const re = /a.b/s;
console.log(re.dotAll);
console.log(re.flags.includes("s"));
"#,
        ["true", "true"]
    };

    regex_sticky_matches_from_current_lastindex => {
        r#"
const re = /foo/y;
re.lastIndex = 4;
const match = re.exec("bar foo foo");
console.log(match[0]);
console.log(re.lastIndex);
"#,
        ["foo", "7"]
    };

    regex_sticky_fails_if_match_does_not_start_at_lastindex => {
        r#"
const re = /foo/y;
re.lastIndex = 0;
console.log(re.exec("bar foo") === null);
console.log(re.lastIndex);
"#,
        ["true", "0"]
    };

    regex_hasindices_reports_full_match_bounds => {
        r#"
const re = /(cat)/d;
const match = re.exec("a cat nap");
console.log(match.indices[0].join(","));
"#,
        ["2,5"]
    };

    regex_hasindices_reports_capture_group_bounds => {
        r#"
const re = /(\d+)-(\d+)/d;
const match = re.exec("date 2024-07");
console.log(match.indices[1].join(","));
console.log(match.indices[2].join(","));
"#,
        ["5,9", "10,12"]
    };

    regex_unicode_sets_union_matches_either_branch => {
        r#"
const re = /[\p{ASCII}--[0-9]]/v;
console.log(re.test("A"));
console.log(re.test("7"));
"#,
        ["true", "false"]
    };

    regex_unicode_sets_string_properties_match_emoji => {
        r#"
const re = /\p{RGI_Emoji}/v;
console.log(re.test("🙂"));
console.log(re.test("A"));
"#,
        ["true", "false"]
    };
}
