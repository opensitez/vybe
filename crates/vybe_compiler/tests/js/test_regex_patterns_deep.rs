/// Regular expression anchors, word boundaries, lookaheads, lookbehinds
use super::helpers::run_js;

#[test]
fn caret_anchors_start() {
    assert_eq!(
        run_js(
            r#"
console.log(/^hello/.test("hello world"));
console.log(/^hello/.test("say hello"));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn dollar_anchors_end() {
    assert_eq!(
        run_js(
            r#"
console.log(/world$/.test("hello world"));
console.log(/world$/.test("world order"));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn word_boundary_b() {
    assert_eq!(
        run_js(
            r#"
const re = /\bcat\b/;
console.log(re.test("cat"));
console.log(re.test("cats"));
console.log(re.test("the cat sat"));
console.log(re.test("concatenate"));
"#
        ),
        vec!["true", "false", "true", "false"]
    );
}

#[test]
fn positive_lookahead() {
    assert_eq!(
        run_js(
            r#"
// Match digits followed by "px"
const re = /\d+(?=px)/;
const m = "width: 100px, height: 200em".match(re);
console.log(m[0]);
"#
        ),
        vec!["100"]
    );
}

#[test]
fn negative_lookahead() {
    assert_eq!(
        run_js(
            r#"
// Match digits NOT followed by "px"
const re = /\d+(?!px)/g;
const matches = [..."200px 300em 400px".matchAll(re)].map(m => m[0]);
console.log(matches.join(","));
"#
        ),
        vec!["20,300,40"]
    );
}

#[test]
fn positive_lookbehind() {
    assert_eq!(
        run_js(
            r#"
// Match digits preceded by "$"
const re = /(?<=\$)\d+/g;
const matches = [..."$100 €200 $300".matchAll(re)].map(m => m[0]);
console.log(matches.join(","));
"#
        ),
        vec!["100,300"]
    );
}

#[test]
fn negative_lookbehind() {
    assert_eq!(
        run_js(
            r#"
// Match digits NOT preceded by "$"
const re = /(?<!\$)\d+/g;
const text = "$100 200 $300 400";
const matches = [...text.matchAll(re)].map(m => m[0]);
// Only 200 and 400 should match
console.log(matches.join(","));
"#
        ),
        vec!["200,400"]
    );
}

#[test]
fn multiline_mode_anchors() {
    assert_eq!(
        run_js(
            r#"
const text = "first\nsecond\nthird";
const matches = text.match(/^\w+/mg);
console.log(matches.join(","));
"#
        ),
        vec!["first,second,third"]
    );
}

#[test]
fn dotall_matches_newline() {
    assert_eq!(
        run_js(
            r#"
const re = /start.+end/s; // s flag = dotAll
const text = "start\nmiddle\nend";
console.log(re.test(text));
// Without s flag
const re2 = /start.+end/;
console.log(re2.test(text));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn alternation_in_group() {
    assert_eq!(
        run_js(
            r#"
const re = /^(cat|dog|bird)$/;
console.log(re.test("cat"));
console.log(re.test("dog"));
console.log(re.test("fish"));
"#
        ),
        vec!["true", "true", "false"]
    );
}

#[test]
fn non_capturing_group() {
    assert_eq!(
        run_js(
            r#"
const re = /(?:foo)(bar)/;
const m = "foobar".match(re);
console.log(m[0]); // full match
console.log(m[1]); // first capturing group (bar)
console.log(m[2]); // no second group
"#
        ),
        vec!["foobar", "bar", "undefined"]
    );
}

#[test]
fn quantifier_greedy_vs_lazy() {
    assert_eq!(
        run_js(
            r#"
const text = "<a><b><c>";
const greedy = text.match(/<.*>/);
const lazy = text.match(/<.*?>/);
console.log(greedy[0]);
console.log(lazy[0]);
"#
        ),
        vec!["<a><b><c>", "<a>"]
    );
}
