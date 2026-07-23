use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: string operations — format, f-strings, encode/decode, methods, slicing, template strings
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_strings_fstring_expressions() {
    let src = r#"
name = "World"
n = 42
print(f"Hello, {name}!")
print(f"{n:08b}")    # binary, zero-padded to 8
print(f"{3.14159:.2f}")
print(f"{name!r}")   # repr
print(f"{name!s}")   # str
print(f"{n:+d}")     # explicit sign
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Hello, World!",
            "00101010",
            "3.14",
            "'World'",
            "World",
            "+42"
        ]
    );
}

#[test]
fn test_py_strings_format_method() {
    let src = r#"
print("{} and {}".format("Alice", "Bob"))
print("{name} is {age}".format(name="Charlie", age=25))
print("{0[key]}".format({"key": "value"}))
print("{:.3f}".format(3.14159))
"#;
    assert_eq!(
        run_python(src),
        vec!["Alice and Bob", "Charlie is 25", "value", "3.142"]
    );
}

#[test]
fn test_py_strings_case_methods() {
    let src = r#"
s = "hello world"
print(s.upper())
print(s.title())
print(s.capitalize())
print("PYTHON".lower())
print("hello World".swapcase())
"#;
    assert_eq!(
        run_python(src),
        vec![
            "HELLO WORLD",
            "Hello World",
            "Hello world",
            "python",
            "HELLO wORLD"
        ]
    );
}

#[test]
fn test_py_strings_strip_split_join() {
    let src = r#"
s = "  hello world  "
print(s.strip())
print(s.lstrip())
print(s.rstrip())
parts = "a,b,c,d".split(",")
print(parts)
print("|".join(parts))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "hello world",
            "hello world  ",
            "  hello world",
            "['a', 'b', 'c', 'd']",
            "a|b|c|d"
        ]
    );
}

#[test]
fn test_py_strings_find_index_contains() {
    let src = r#"
s = "Hello, World!"
print(s.find("World"))
print(s.find("xyz"))
print(s.index("World"))
try:
    s.index("xyz")
except ValueError:
    print("not found")
print("World" in s)
"#;
    assert_eq!(run_python(src), vec!["7", "-1", "7", "not found", "True"]);
}

#[test]
fn test_py_strings_replace_count() {
    let src = r#"
s = "aabbccaabb"
print(s.replace("aa", "XX"))
print(s.replace("aa", "XX", 1))
print(s.count("aa"))
print(s.count("a"))
"#;
    assert_eq!(run_python(src), vec!["XXbbccXXbb", "XXbbccaabb", "2", "4"]);
}

#[test]
fn test_py_strings_startswith_endswith() {
    let src = r#"
url = "https://example.com/path"
print(url.startswith("https://"))
print(url.startswith(("http://", "https://")))
print(url.endswith(".com/path"))
print(url.endswith((".json", ".html", "/path")))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_strings_encode_decode() {
    let src = r#"
text = "café"
encoded = text.encode("utf-8")
print(type(encoded).__name__)
print(len(encoded))
print(encoded.decode("utf-8"))
print(text.encode("ascii", errors="replace"))
"#;
    assert_eq!(run_python(src), vec!["bytes", "5", "café", "b'caf?'"]);
}

#[test]
fn test_py_strings_slicing_and_reversal() {
    let src = r#"
s = "Hello, World!"
print(s[7:12])
print(s[:5])
print(s[-6:-1])
print(s[::-1])   # reverse
print(s[::2])    # every other character
"#;
    assert_eq!(
        run_python(src),
        vec!["World", "Hello", "World", "!dlroW ,olleH", "Hlo ol!"]
    );
}

#[test]
fn test_py_strings_partition_and_rpartition() {
    let src = r#"
path = "/usr/local/bin/python"
before, sep, after = path.partition("/")
print(repr(before), sep, after)

before2, sep2, after2 = path.rpartition("/")
print(before2, sep2, after2)
"#;
    assert_eq!(
        run_python(src),
        vec!["'' / usr/local/bin/python", "/usr/local/bin / python"]
    );
}

#[test]
fn test_py_strings_zfill_center_ljust_rjust() {
    let src = r#"
print("42".zfill(6))
print("hi".center(10, "*"))
print("left".ljust(10, "."))
print("right".rjust(10, "."))
"#;
    assert_eq!(
        run_python(src),
        vec!["000042", "****hi****", "left......", ".....right"]
    );
}

#[test]
fn test_py_strings_isdigit_isalpha_checks() {
    let src = r#"
print("123".isdigit())
print("abc".isalpha())
print("abc123".isalnum())
print("  ".isspace())
print("Hello World".istitle())
print("HELLO".isupper())
print("hello".islower())
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "True", "True", "True", "True", "True", "True"]
    );
}

#[test]
fn test_py_strings_splitlines() {
    let src = r#"
text = "line1\nline2\r\nline3\rline4"
lines = text.splitlines()
print(len(lines))
print(lines)
"#;
    assert_eq!(
        run_python(src),
        vec!["4", "['line1', 'line2', 'line3', 'line4']"]
    );
}

#[test]
fn test_py_strings_translate_maketrans() {
    let src = r#"
table = str.maketrans("aeiou", "AEIOU", " ")
result = "hello world".translate(table)
print(result)
"#;
    assert_eq!(run_python(src), vec!["hEllOwOrld"]);
}

#[test]
fn test_py_strings_template_string() {
    let src = r#"
from string import Template

t = Template("Hello, $name! You have $count messages.")
print(t.substitute(name="Alice", count=5))
print(t.safe_substitute(name="Bob"))  # missing keys left as-is
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Hello, Alice! You have 5 messages.",
            "Hello, Bob! You have $count messages."
        ]
    );
}

#[test]
fn test_py_strings_format_spec_alignment() {
    let src = r#"
print(f"{'left':<10}|")
print(f"{'center':^10}|")
print(f"{'right':>10}|")
print(f"{42:010d}")
print(f"{255:#010x}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "left      |",
            "  center  |",
            "     right|",
            "0000000042",
            "0x000000ff"
        ]
    );
}

#[test]
fn test_py_strings_multiline_and_raw() {
    let src = r#"
ml = """line1
line2
line3"""
print(len(ml.splitlines()))

raw = r"C:\Users\name\file.txt"
print(raw)
print(raw.count("\\"))
"#;
    assert_eq!(run_python(src), vec!["3", r"C:\Users\name\file.txt", "3"]);
}
