use super::helpers::run_python;

// tomllib — loads, load, TOML data types (table, inline table, array, array of tables, datetime, boolean, float, integer, string)

#[test]
fn test_tomllib_loads_basic_key_value() {
    let out = run_python(r#"
import tomllib
toml_str = """
title = "TOML Example"
count = 42
pi = 3.14159
enabled = true
"""
d = tomllib.loads(toml_str)
print(d["title"])
print(d["count"])
print(d["pi"])
print(d["enabled"])
"#);
    assert_eq!(out, vec!["TOML Example", "42", "3.14159", "True"]);
}

#[test]
fn test_tomllib_loads_tables() {
    let out = run_python(r#"
import tomllib
toml_str = """
[owner]
name = "Alice"
dob = 1979-05-27T07:32:00Z

[database]
server = "192.168.1.1"
ports = [ 8001, 8002, 8003 ]
connection_max = 5000
"""
d = tomllib.loads(toml_str)
print(d["owner"]["name"])
print(d["database"]["server"])
print(d["database"]["ports"])
"#);
    assert_eq!(out, vec!["Alice", "192.168.1.1", "[8001, 8002, 8003]"]);
}

#[test]
fn test_tomllib_loads_array_of_tables() {
    let out = run_python(r#"
import tomllib
toml_str = """
[[products]]
name = "Hammer"
sku = 738592

[[products]]
name = "Nail"
sku = 284758
color = "gray"
"""
d = tomllib.loads(toml_str)
print(len(d["products"]))
print(d["products"][0]["name"])
print(d["products"][1]["color"])
"#);
    assert_eq!(out, vec!["2", "Hammer", "gray"]);
}

#[test]
fn test_tomllib_loads_inline_tables() {
    let out = run_python(r#"
import tomllib
toml_str = """
name = { first = "Tom", last = "Preston-Werner" }
point = { x = 1, y = 2 }
"""
d = tomllib.loads(toml_str)
print(d["name"]["first"])
print(d["point"]["x"], d["point"]["y"])
"#);
    assert_eq!(out, vec!["Tom", "1 2"]);
}

#[test]
fn test_tomllib_loads_datetime_parsing() {
    let out = run_python(r#"
import tomllib
from datetime import datetime, date, time
toml_str = """
odt = 1979-05-27T07:32:00Z
ldt = 1979-05-27T07:32:00
ld  = 1979-05-27
lt  = 07:32:00
"""
d = tomllib.loads(toml_str)
print(isinstance(d["odt"], datetime))
print(isinstance(d["ldt"], datetime))
print(isinstance(d["ld"], date))
print(isinstance(d["lt"], time))
"#);
    assert_eq!(out, vec!["True", "True", "True", "True"]);
}

#[test]
fn test_tomllib_loads_multiline_basic_string() {
    let out = run_python(r#"
import tomllib
toml_str = '''
str = """
The quick brown \\
fox jumps over \\
the lazy dog."""
'''
d = tomllib.loads(toml_str)
print(d["str"].strip())
"#);
    assert_eq!(out, vec!["The quick brown fox jumps over the lazy dog."]);
}

#[test]
fn test_tomllib_loads_literal_strings() {
    let out = run_python(r#"
import tomllib
toml_str = '''
winpath  = 'C:\\Users\\nodejs\\templates'
winpath2 = '\\\\ServerX\\admin$\\system32\\'
quoted   = 'Tom "Dhne" Preston-Werner'
regex    = '<\\i\\c*\\s*>'
'''
d = tomllib.loads(toml_str)
print(d["winpath"])
print(d["quoted"])
"#);
    assert_eq!(out, vec!["C:\\Users\\nodejs\\templates", "Tom \"Dhne\" Preston-Werner"]);
}

#[test]
fn test_tomllib_loads_integer_formats() {
    let out = run_python(r#"
import tomllib
toml_str = """
int1 = +99
int2 = 42
int3 = 0
int4 = -17
hex1 = 0xDEADBEEF
oct1 = 0o755
bin1 = 0b11010110
"""
d = tomllib.loads(toml_str)
print(d["int1"])
print(d["hex1"])
print(d["oct1"])
print(d["bin1"])
"#);
    assert_eq!(out, vec!["99", "3735928559", "493", "214"]);
}

#[test]
fn test_tomllib_loads_float_formats() {
    let out = run_python(r#"
import tomllib
toml_str = """
flt1 = +1.0
flt2 = 3.1415
flt3 = -0.01
flt4 = 5e+22
flt5 = 1e06
flt6 = -2E-2
"""
d = tomllib.loads(toml_str)
print(d["flt1"])
print(d["flt5"])
print(d["flt6"])
"#);
    assert_eq!(out, vec!["1.0", "1000000.0", "-0.02"]);
}

#[test]
fn test_tomllib_loads_special_floats() {
    let out = run_python(r#"
import tomllib, math
toml_str = """
sf1 = inf
sf2 = +inf
sf3 = -inf
sf4 = nan
"""
d = tomllib.loads(toml_str)
print(d["sf1"] == float("inf"))
print(d["sf3"] == float("-inf"))
print(math.isnan(d["sf4"]))
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_tomllib_loads_parse_float_custom() {
    let out = run_python(r#"
import tomllib
from decimal import Decimal
toml_str = "val = 3.14"
d = tomllib.loads(toml_str, parse_float=Decimal)
print(type(d["val"]).__name__)
print(d["val"])
"#);
    assert_eq!(out, vec!["Decimal", "3.14"]);
}

#[test]
fn test_tomllib_load_bytes_stream() {
    let out = run_python(r#"
import tomllib, io
b = b"key = 'value'\n"
d = tomllib.load(io.BytesIO(b))
print(d["key"])
"#);
    assert_eq!(out, vec!["value"]);
}

#[test]
fn test_tomllib_invalid_syntax_raises_decode_error() {
    let out = run_python(r#"
import tomllib
try:
    tomllib.loads("invalid = [ unclosed array")
except tomllib.TOMLDecodeError:
    print("TOMLDecodeError")
"#);
    assert_eq!(out, vec!["TOMLDecodeError"]);
}

#[test]
fn test_tomllib_dotted_keys() {
    let out = run_python(r#"
import tomllib
toml_str = """
physical.color = "orange"
physical.shape = "round"
site.google.com = true
"""
d = tomllib.loads(toml_str)
print(d["physical"]["color"])
print(d["site"]["google"]["com"])
"#);
    assert_eq!(out, vec!["orange", "True"]);
}

#[test]
fn test_tomllib_nested_tables() {
    let out = run_python(r#"
import tomllib
toml_str = """
[a.b.c]
d = "hello"
"""
d = tomllib.loads(toml_str)
print(d["a"]["b"]["c"]["d"])
"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn test_tomllib_numeric_underscores() {
    let out = run_python(r#"
import tomllib
toml_str = """
million = 1_000_000
float_val = 3.141_592
"""
d = tomllib.loads(toml_str)
print(d["million"])
print(d["float_val"])
"#);
    assert_eq!(out, vec!["1000000", "3.141592"]);
}

#[test]
fn test_tomllib_empty_toml_document() {
    let out = run_python(r#"
import tomllib
d = tomllib.loads("")
print(d)
"#);
    assert_eq!(out, vec!["{}"]);
}

#[test]
fn test_tomllib_comment_only_document() {
    let out = run_python(r##"
import tomllib
d = tomllib.loads("# This is a comment\n# Another comment\n")
print(d)
"##);
    assert_eq!(out, vec!["{}"]);
}

#[test]
fn test_tomllib_duplicate_keys_raises_decode_error() {
    let out = run_python(r#"
import tomllib
toml_str = """
dupe = 1
dupe = 2
"""
try:
    tomllib.loads(toml_str)
except tomllib.TOMLDecodeError:
    print("TOMLDecodeError")
"#);
    assert_eq!(out, vec!["TOMLDecodeError"]);
}

#[test]
fn test_tomllib_nested_arrays() {
    let out = run_python(r#"
import tomllib
toml_str = "matrix = [[1, 2], [3, 4]]"
d = tomllib.loads(toml_str)
print(d["matrix"])
"#);
    assert_eq!(out, vec!["[[1, 2], [3, 4]]"]);
}
