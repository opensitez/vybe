use super::helpers::run_python;

// token + tokenize — generate_tokens, TokenInfo fields, token type constants

#[test]
fn test_tokenize_basic_expression() {
    let out = run_python(r#"
import tokenize, io, token
src = "x = 1 + 2\n"
tokens = list(tokenize.generate_tokens(io.StringIO(src).readline))
names = [tok.string for tok in tokens if tok.type not in (token.ENCODING, token.ENDMARKER, token.NEWLINE, token.NL)]
print(names)
"#);
    assert_eq!(out, vec!["['x', '=', '1', '+', '2']"]);
}

#[test]
fn test_tokenize_tokeninfo_type_field() {
    let out = run_python(r#"
import tokenize, io, token
src = "hello\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
names_tok = [t for t in toks if t.type == token.NAME]
print(len(names_tok) == 1)
print(names_tok[0].string)
"#);
    assert_eq!(out, vec!["True", "hello"]);
}

#[test]
fn test_tokenize_tokeninfo_string_field() {
    let out = run_python(r#"
import tokenize, io
src = '"hello world"\n'
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
strings = [t.string for t in toks if t.type == tokenize.STRING]
print(strings)
"#);
    assert_eq!(out, vec!["['\"hello world\"']"]);
}

#[test]
fn test_tokenize_number_tokens() {
    let out = run_python(r#"
import tokenize, io, token
src = "42 3.14 0xFF\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
numbers = [t.string for t in toks if t.type == token.NUMBER]
print(numbers)
"#);
    assert_eq!(out, vec!["['42', '3.14', '0xFF']"]);
}

#[test]
fn test_tokenize_operator_tokens() {
    let out = run_python(r#"
import tokenize, io, token
src = "a + b * c\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
ops = [t.string for t in toks if t.type == token.OP]
print(ops)
"#);
    assert_eq!(out, vec!["['+', '*']"]);
}

#[test]
fn test_tokenize_comment_token() {
    let out = run_python(r#"
import tokenize, io
src = "x = 1  # this is a comment\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
comments = [t.string for t in toks if t.type == tokenize.COMMENT]
print(comments)
"#);
    assert_eq!(out, vec!["['# this is a comment']"]);
}

#[test]
fn test_tokenize_tokeninfo_start_end() {
    let out = run_python(r#"
import tokenize, io, token
src = "x = 42\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
num = [t for t in toks if t.type == token.NUMBER][0]
print(num.start[0])  # line 1
print(num.end[0])    # same line
print(num.start[1] < num.end[1])  # col start < col end
"#);
    assert_eq!(out, vec!["1", "1", "True"]);
}

#[test]
fn test_tokenize_tokeninfo_line_field() {
    let out = run_python(r#"
import tokenize, io, token
src = "answer = 42\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
num = [t for t in toks if t.type == token.NUMBER][0]
print("42" in num.line)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tokenize_multiline_source() {
    let out = run_python(r#"
import tokenize, io, token
src = "a = 1\nb = 2\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
names = [t.string for t in toks if t.type == token.NAME]
print(sorted(names))
"#);
    assert_eq!(out, vec!["['a', 'b']"]);
}

#[test]
fn test_tokenize_indent_dedent_tokens() {
    let out = run_python(r#"
import tokenize, io
src = "if True:\n    pass\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
types = [t.type for t in toks]
print(tokenize.INDENT in types)
print(tokenize.DEDENT in types)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_token_constants_values() {
    let out = run_python(r#"
import token
print(token.NAME > 0)
print(token.NUMBER > 0)
print(token.STRING > 0)
print(token.OP > 0)
"#);
    assert_eq!(out, vec!["True", "True", "True", "True"]);
}

#[test]
fn test_token_tok_name_lookup() {
    let out = run_python(r#"
import token
print(token.tok_name[token.NAME])
print(token.tok_name[token.NUMBER])
print(token.tok_name[token.OP])
"#);
    assert_eq!(out, vec!["NAME", "NUMBER", "OP"]);
}

#[test]
fn test_tokenize_string_with_backslash() {
    let out = run_python(r#"
import tokenize, io, token
src = r'"hello\nworld"' + "\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
strings = [t.string for t in toks if t.type == token.STRING]
print(len(strings) == 1)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tokenize_fstring_token() {
    let out = run_python(r#"
import tokenize, io, token, sys
src = 'f"hello {name}"\n'
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
# f-strings may be STRING or multiple tokens depending on Python version
string_toks = [t for t in toks if t.type == token.STRING or t.type == token.NAME]
print(len(string_toks) >= 1)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tokenize_detect_encoding_utf8() {
    let out = run_python(r##"
import tokenize, io
src = b"# -*- coding: utf-8 -*-\nx = 1\n"
toks = list(tokenize.tokenize(io.BytesIO(src).readline))
encodings = [t.string for t in toks if t.type == tokenize.ENCODING]
print(encodings)
"##);
    assert_eq!(out, vec!["['utf-8']"]);
}

#[test]
fn test_tokenize_untokenize_roundtrip() {
    let out = run_python(r#"
import tokenize, io
src = "x = 1 + 2\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
result = tokenize.untokenize(toks)
print("x" in result)
print("1" in result)
print("+" in result)
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_tokenize_keyword_token_type() {
    let out = run_python(r#"
import tokenize, io, token
src = "if True:\n    pass\n"
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
names = [t.string for t in toks if t.type == token.NAME]
print("if" in names)
print("True" in names)
print("pass" in names)
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_tokenize_empty_source() {
    let out = run_python(r#"
import tokenize, io
src = ""
toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
types = [t.type for t in toks]
print(tokenize.ENDMARKER in types)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_tokenize_error_on_invalid_token() {
    let out = run_python(r#"
import tokenize, io
src = "$\n"
try:
    list(tokenize.generate_tokens(io.StringIO(src).readline))
    print("no error")
except tokenize.TokenError:
    print("TokenError")
except Exception:
    print("other error")
"#);
    // $ is an error token in Python
    assert_eq!(out, vec!["other error"]);
}

#[test]
fn test_token_iseof() {
    let out = run_python(r#"
import token
print(token.ENDMARKER == 0)
"#);
    assert_eq!(out, vec!["True"]);
}
