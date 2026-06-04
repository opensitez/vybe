/// Regex advanced patterns — lookaheads, lookbehinds, named groups, flags
use super::helpers::run_js;

#[test]
fn named_capture_groups_replace() {
    assert_eq!(
        run_js(
            r#"
const date = "2024-06-15";
const reordered = date.replace(
    /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/,
    "$<day>/$<month>/$<year>"
);
console.log(reordered);
"#
        ),
        vec!["15/06/2024"]
    );
}

#[test]
fn positive_lookbehind() {
    assert_eq!(
        run_js(
            r#"
const prices = "apple: $10, banana: $5, cherry: $15";
const amounts = prices.match(/(?<=\$)\d+/g);
console.log(amounts.join(","));
"#
        ),
        vec!["10,5,15"]
    );
}

#[test]
fn negative_lookahead_pattern() {
    assert_eq!(
        run_js(
            r#"
// Match words not followed by a comma
const text = "hello, world, foo bar baz";
const words = text.match(/\b\w+\b(?!,)/g);
console.log(words.join(","));
"#
        ),
        vec!["world,foo,bar,baz"]
    );
}

#[test]
fn regex_global_replace_function() {
    assert_eq!(
        run_js(
            r#"
const result = "hello world".replace(/(\w+)/g, (match, word) => word[0].toUpperCase() + word.slice(1));
console.log(result);
"#
        ),
        vec!["Hello World"]
    );
}

#[test]
fn regex_split_with_capture() {
    assert_eq!(
        run_js(
            r#"
const parts = "one1two22three333four".split(/(\d+)/);
console.log(parts.join("|"));
"#
        ),
        vec!["one|1|two|22|three|333|four"]
    );
}

#[test]
fn regex_match_all_groups() {
    assert_eq!(
        run_js(
            r#"
const html = '<a href="http://foo.com">Foo</a> <a href="http://bar.com">Bar</a>';
const re = /<a href="([^"]+)">([^<]+)<\/a>/g;
const links = [...html.matchAll(re)].map(m => `${m[2]}:${m[1]}`);
console.log(links.join("|"));
"#
        ),
        vec!["Foo:http://foo.com|Bar:http://bar.com"]
    );
}

#[test]
fn regex_unicode_flag() {
    assert_eq!(
        run_js(
            r#"
const emoji = "Hello 😀 World 🌍";
const emojiCount = (emoji.match(/\p{Emoji}/gu) || []).length;
console.log(emojiCount >= 2);
const wordCount = "hello world".match(/\p{L}+/gu).length;
console.log(wordCount);
"#
        ),
        vec!["true", "2"]
    );
}

#[test]
fn regex_dotall_flag() {
    assert_eq!(
        run_js(
            r#"
const text = "line1\nline2\nline3";
const withDot = text.match(/line1.line2/s);
const withoutDot = text.match(/line1.line2/);
console.log(withDot !== null);
console.log(withoutDot);
"#
        ),
        vec!["true", "null"]
    );
}

#[test]
fn regex_sticky_flag() {
    assert_eq!(
        run_js(
            r#"
const re = /\d+/y;
re.lastIndex = 3;
const str = "abc123def456";
const m1 = re.exec(str);
console.log(m1[0]);
console.log(re.lastIndex);
const m2 = re.exec(str);
console.log(m2);
"#
        ),
        vec!["123", "6", "null"]
    );
}

#[test]
fn regex_backreference() {
    assert_eq!(
        run_js(
            r#"
// Match doubled words
const doubled = /\b(\w+) \1\b/g;
const text = "the the quick brown fox fox";
const matches = text.match(doubled);
console.log(matches.join(","));
"#
        ),
        vec!["the the,fox fox"]
    );
}

#[test]
fn regex_quantifiers_greedy_lazy() {
    assert_eq!(
        run_js(
            r#"
const html = "<b>bold</b> and <i>italic</i>";
const greedy = html.match(/<.+>/)[0];
const lazy = html.match(/<.+?>/)[0];
console.log(greedy);
console.log(lazy);
"#
        ),
        vec!["<b>bold</b> and <i>italic</i>", "<b>"]
    );
}

#[test]
fn regex_assertions() {
    assert_eq!(
        run_js(
            r#"
const wordStart = "cat concatenate catalog".match(/\bcat\b/g);
console.log(wordStart.length);
const atLineEnd = "hello\nworld".match(/hello$/m);
console.log(atLineEnd !== null);
"#
        ),
        vec!["1", "true"]
    );
}
