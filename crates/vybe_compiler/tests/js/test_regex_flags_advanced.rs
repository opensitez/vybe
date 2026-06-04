/// Regex /d flag (indices), sticky /y flag, exec loop, matchAll patterns
use super::helpers::run_js;

#[test]
fn sticky_flag_anchors_at_last_index() {
    assert_eq!(
        run_js(
            r#"
const re = /\d+/y;
re.lastIndex = 4;
const m = re.exec("abc 123 def");
console.log(m ? m[0] : null);
console.log(re.lastIndex);
"#
        ),
        vec!["123", "7"]
    );
}

#[test]
fn sticky_no_match_resets_lastindex() {
    assert_eq!(
        run_js(
            r#"
const re = /\d+/y;
re.lastIndex = 0;
const m = re.exec("abc");
console.log(m);
console.log(re.lastIndex);
"#
        ),
        vec!["null", "0"]
    );
}

#[test]
fn exec_global_loop_collects_all() {
    assert_eq!(
        run_js(
            r#"
const re = /(\w+)=(\d+)/g;
const str = "a=1 b=2 c=3";
const matches = [];
let m;
while ((m = re.exec(str)) !== null) {
    matches.push(m[1] + ":" + m[2]);
}
console.log(matches.join(","));
"#
        ),
        vec!["a:1,b:2,c:3"]
    );
}

#[test]
fn exec_captures_named_groups() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
const m = re.exec("2024-06-15");
console.log(m.groups.year);
console.log(m.groups.month);
console.log(m.groups.day);
"#
        ),
        vec!["2024", "06", "15"]
    );
}

#[test]
fn match_all_returns_all_with_groups() {
    assert_eq!(
        run_js(
            r#"
const re = /(\w+)=(\d+)/g;
const results = [...("a=1 b=2".matchAll(re))];
console.log(results.length);
console.log(results[0][1]);
console.log(results[1][2]);
"#
        ),
        vec!["2", "a", "2"]
    );
}

#[test]
fn match_all_includes_index() {
    assert_eq!(
        run_js(
            r#"
const re = /cat/g;
const matches = [...("catfish cat caterpillar".matchAll(re))];
console.log(matches.map(m => m.index).join(","));
"#
        ),
        vec!["0,8,12"]
    );
}

#[test]
fn regex_d_flag_gives_indices() {
    assert_eq!(
        run_js(
            r#"
const re = /(?<name>\w+)/d;
const m = re.exec("hello world");
console.log(m.indices[0][0]); // start of full match
console.log(m.indices[0][1]); // end
console.log(m.indices.groups.name[0]); // group start
"#
        ),
        vec!["0", "5", "0"]
    );
}

#[test]
fn regex_d_flag_capture_group_indices() {
    assert_eq!(
        run_js(
            r#"
const re = /(\d{4})-(\d{2})/d;
const m = re.exec("Date: 2024-06");
console.log(m.indices[1].join(",")); // year capture indices
console.log(m.indices[2].join(",")); // month capture
"#
        ),
        vec!["6,10", "11,13"]
    );
}

#[test]
fn test_flag_checks() {
    assert_eq!(
        run_js(
            r#"
const re = /abc/gim;
console.log(re.global);
console.log(re.ignoreCase);
console.log(re.multiline);
console.log(re.flags.split("").sort().join(""));
"#
        ),
        vec!["true", "true", "true", "gim"]
    );
}

#[test]
fn regex_source_and_flags_properties() {
    assert_eq!(
        run_js(
            r#"
const re = /hello\s+world/gi;
console.log(re.source);
console.log(re.flags);
"#
        ),
        vec!["hello\\s+world", "gi"]
    );
}

#[test]
fn regex_lastindex_manual_control() {
    assert_eq!(
        run_js(
            r#"
const re = /a/g;
const str = "ababa";
re.lastIndex = 2;
const m = re.exec(str);
console.log(m.index); // finds 'a' at index 2
re.lastIndex = 0;
const m2 = re.exec(str);
console.log(m2.index); // starts from beginning
"#
        ),
        vec!["2", "0"]
    );
}

#[test]
fn sticky_and_global_differ() {
    assert_eq!(
        run_js(
            r#"
const sticky = /\w+/y;
const global = /\w+/g;
const str = "  foo bar";
// sticky: must match at lastIndex (0), space is not \w
const ms = sticky.exec(str);
console.log(ms); // null, no match at pos 0
// global: searches anywhere
const mg = global.exec(str);
console.log(mg[0]); // "foo"
"#
        ),
        vec!["null", "foo"]
    );
}
