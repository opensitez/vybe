use super::helpers::run_python;

// html.parser — HTMLParser, feed, handle_starttag/endtag/data
// html — unescape entity references

#[test]
fn test_html_unescape_named_entity() {
    let out = run_python(
        r#"
import html
print(html.unescape("&amp;"))
print(html.unescape("&lt;"))
print(html.unescape("&gt;"))
"#,
    );
    assert_eq!(out, vec!["&", "<", ">"]);
}

#[test]
fn test_html_unescape_numeric_decimal_entity() {
    let out = run_python(
        r#"
import html
print(html.unescape("&#65;"))
"#,
    );
    assert_eq!(out, vec!["A"]);
}

#[test]
fn test_html_unescape_numeric_hex_entity() {
    let out = run_python(
        r#"
import html
print(html.unescape("&#x41;"))
"#,
    );
    assert_eq!(out, vec!["A"]);
}

#[test]
fn test_html_unescape_nbsp() {
    let out = run_python(
        r#"
import html
result = html.unescape("a&nbsp;b")
print(result)
"#,
    );
    assert_eq!(out, vec!["a\u{a0}b"]);
}

#[test]
fn test_html_escape_special_chars() {
    let out = run_python(
        r#"
import html
print(html.escape("<script>alert('xss')</script>"))
"#,
    );
    assert_eq!(
        out,
        vec!["&lt;script&gt;alert(&#x27;xss&#x27;)&lt;/script&gt;"]
    );
}

#[test]
fn test_html_escape_with_quote_false() {
    let out = run_python(
        r#"
import html
print(html.escape('"hello"', quote=False))
"#,
    );
    assert_eq!(out, vec!["\"hello\""]);
}

#[test]
fn test_htmlparser_starttag() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_starttag(self, tag, attrs):
        print(tag, sorted(attrs))
p = P()
p.feed('<a href="http://example.com" class="link">')
"#,
    );
    assert_eq!(
        out,
        vec!["a [('class', 'link'), ('href', 'http://example.com')]"]
    );
}

#[test]
fn test_htmlparser_endtag() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_endtag(self, tag):
        print(tag)
p = P()
p.feed("</div>")
"#,
    );
    assert_eq!(out, vec!["div"]);
}

#[test]
fn test_htmlparser_data() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_data(self, data):
        print(repr(data))
p = P()
p.feed("<p>Hello World</p>")
"#,
    );
    assert_eq!(out, vec!["'Hello World'"]);
}

#[test]
fn test_htmlparser_comment() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_comment(self, data):
        print(data.strip())
p = P()
p.feed("<!-- this is a comment -->")
"#,
    );
    assert_eq!(out, vec!["this is a comment"]);
}

#[test]
fn test_htmlparser_self_closing_tag() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_startendtag(self, tag, attrs):
        print(f"selfclose:{tag}")
p = P()
p.feed("<br/>")
"#,
    );
    assert_eq!(out, vec!["selfclose:br"]);
}

#[test]
fn test_htmlparser_multiple_tags() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
tags = []
class P(HTMLParser):
    def handle_starttag(self, tag, attrs):
        tags.append(tag)
p = P()
p.feed("<html><head></head><body></body></html>")
print(tags)
"#,
    );
    assert_eq!(out, vec!["['html', 'head', 'body']"]);
}

#[test]
fn test_htmlparser_attr_with_no_value() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_starttag(self, tag, attrs):
        print(attrs)
p = P()
p.feed("<input disabled>")
"#,
    );
    assert_eq!(out, vec!["[('disabled', None)]"]);
}

#[test]
fn test_htmlparser_decl() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_decl(self, decl):
        print(decl)
p = P()
p.feed("<!DOCTYPE html>")
"#,
    );
    assert_eq!(out, vec!["DOCTYPE html"]);
}

#[test]
fn test_htmlparser_getpos_tracks_line() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_starttag(self, tag, attrs):
        line, col = self.getpos()
        print(line >= 1)
p = P()
p.feed("<div>")
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_htmlparser_nested_content_order() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
events = []
class P(HTMLParser):
    def handle_starttag(self, tag, attrs): events.append(("start", tag))
    def handle_endtag(self, tag): events.append(("end", tag))
    def handle_data(self, d):
        if d.strip(): events.append(("data", d.strip()))
p = P()
p.feed("<p>text</p>")
print(events)
"#,
    );
    assert_eq!(
        out,
        vec!["[('start', 'p'), ('data', 'text'), ('end', 'p')]"]
    );
}

#[test]
fn test_html_unescape_double_amp() {
    let out = run_python(
        r#"
import html
print(html.unescape("&amp;amp;"))
"#,
    );
    assert_eq!(out, vec!["&amp;"]);
}

#[test]
fn test_htmlparser_reset_clears_state() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
p = HTMLParser()
p.feed("<div>")
p.reset()
# After reset we can feed again without error
p.feed("<span>")
print("ok")
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_html_escape_ampersand() {
    let out = run_python(
        r#"
import html
print(html.escape("a & b"))
"#,
    );
    assert_eq!(out, vec!["a &amp; b"]);
}

#[test]
fn test_htmlparser_pi_instruction() {
    let out = run_python(
        r#"
from html.parser import HTMLParser
class P(HTMLParser):
    def handle_pi(self, data):
        print(data.strip())
p = P()
p.feed("<?xml version='1.0'?>")
"#,
    );
    assert_eq!(out, vec!["xml version='1.0'"]);
}
