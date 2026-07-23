use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: JSON Serialization Custom — json.dumps, loads, JSONEncoder subclass, object_hook, object_pairs_hook, formatting
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_json_custom_encoder_subclass() {
    let src = r#"
import json
from datetime import datetime

class CustomEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, datetime):
            return obj.isoformat()
        return super().default(obj)

data = {"event": "login", "time": datetime(2024, 5, 12, 10, 0, 0)}
serialized = json.dumps(data, cls=CustomEncoder)
print("2024-05-12T10:00:00" in serialized)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_json_object_hook_deserialization() {
    let src = r#"
import json

class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def __repr__(self):
        return f"Point({self.x}, {self.y})"

def point_decoder(dct):
    if "__point__" in dct:
        return Point(dct["x"], dct["y"])
    return dct

json_str = '{"__point__": true, "x": 10, "y": 20}'
p = json.loads(json_str, object_hook=point_decoder)
print(type(p).__name__)
print(p)
"#;
    assert_eq!(run_python(src), vec!["Point", "Point(10, 20)"]);
}

#[test]
fn test_py_json_object_pairs_hook_ordered_parsing() {
    let src = r#"
import json
from collections import OrderedDict

json_str = '{"b": 2, "a": 1, "c": 3}'
data = json.loads(json_str, object_pairs_hook=OrderedDict)
print(type(data).__name__)
print(list(data.keys()))
"#;
    assert_eq!(run_python(src), vec!["OrderedDict", "['b', 'a', 'c']"]);
}

#[test]
fn test_py_json_dumps_indent_separators_sort_keys() {
    let src = r#"
import json

data = {"c": 3, "a": 1, "b": 2}
output = json.dumps(data, sort_keys=True, indent=2)
lines = output.splitlines()
print(lines[1].strip())
print(lines[2].strip())
print(lines[3].strip())
"#;
    assert_eq!(run_python(src), vec!["\"a\": 1,", "\"b\": 2,", "\"c\": 3"]);
}

#[test]
fn test_py_json_dump_load_file_handle() {
    let src = r#"
import json, tempfile, os

data = {"items": [1, 2, 3], "status": "ok"}

with tempfile.NamedTemporaryFile(mode="w+", delete=False) as f:
    fname = f.name
    json.dump(data, f)

with open(fname, "r") as f:
    loaded = json.load(f)

os.unlink(fname)
print(loaded["items"])
print(loaded["status"])
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]", "ok"]);
}

#[test]
fn test_py_json_default_function_arg() {
    let src = r#"
import json

def fallback(obj):
    if isinstance(obj, set):
        return sorted(list(obj))
    raise TypeError(f"Unserializable: {type(obj)}")

data = {"tags": {"web", "api"}}
serialized = json.dumps(data, default=fallback)
print(serialized)
"#;
    assert_eq!(run_python(src), vec!["{\"tags\": [\"api\", \"web\"]}"]);
}

#[test]
fn test_py_json_decode_error_position_info() {
    let src = r#"
import json

invalid_json = '{"key": value}'
try:
    json.loads(invalid_json)
except json.JSONDecodeError as e:
    print(f"JSONDecodeError at line {e.lineno} col {e.colno}")
"#;
    assert_eq!(run_python(src), vec!["JSONDecodeError at line 1 col 9"]);
}

#[test]
fn test_py_json_allow_nan_options() {
    let src = r#"
import json

data = {"val": float("nan")}
s = json.dumps(data, allow_nan=True)
print(s)

try:
    json.dumps(data, allow_nan=False)
except ValueError:
    print("ValueError: Out of range float values are not JSON compliant")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "{\"val\": NaN}",
            "ValueError: Out of range float values are not JSON compliant"
        ]
    );
}

#[test]
fn test_py_json_primitive_types_roundtrip() {
    let src = r#"
import json

primitives = [None, True, False, 100, 3.14, "hello string", [1, 2], {"k": "v"}]
for p in primitives:
    s = json.dumps(p)
    restored = json.loads(s)
    print(restored == p)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "True", "True", "True", "True", "True", "True", "True", "True"
        ]
    );
}

#[test]
fn test_py_json_non_string_dict_keys_coercion() {
    let src = r#"
import json

data = {1: "one", 2: "two"}
s = json.dumps(data)
print(s)
restored = json.loads(s)
print(list(restored.keys()))  # keys coerced to strings in JSON!
"#;
    assert_eq!(
        run_python(src),
        vec!["{\"1\": \"one\", \"2\": \"two\"}", "['1', '2']"]
    );
}
