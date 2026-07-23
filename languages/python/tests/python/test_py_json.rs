use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: json — dumps, loads, encoder, decoder, indent, sort_keys, default, object_hook
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_json_basic_serialize_deserialize() {
    let src = r#"
import json

data = {"name": "Alice", "age": 30, "active": True, "score": None}
serialized = json.dumps(data)
parsed = json.loads(serialized)
print(parsed["name"])
print(parsed["age"])
print(parsed["active"])
print(parsed["score"] is None)
"#;
    assert_eq!(run_python(src), vec!["Alice", "30", "True", "True"]);
}

#[test]
fn test_py_json_dumps_formatting() {
    let src = r#"
import json

data = {"b": 2, "a": 1}
print(json.dumps(data, sort_keys=True))
print(json.dumps(data, sort_keys=True, indent=2))
"#;
    assert_eq!(
        run_python(src),
        vec![r#"{"a": 1, "b": 2}"#, "{\n  \"a\": 1,\n  \"b\": 2\n}"]
    );
}

#[test]
fn test_py_json_custom_encoder() {
    let src = r#"
import json
from datetime import datetime

class DateEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, datetime):
            return obj.isoformat()
        return super().default(obj)

dt = datetime(2024, 1, 15, 10, 30)
result = json.dumps({"ts": dt}, cls=DateEncoder)
print(result)
"#;
    assert_eq!(run_python(src), vec![r#"{"ts": "2024-01-15T10:30:00"}"#]);
}

#[test]
fn test_py_json_dumps_default_function() {
    let src = r#"
import json

class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

def point_encoder(obj):
    if isinstance(obj, Point):
        return {"x": obj.x, "y": obj.y}
    raise TypeError(f"Not serializable: {type(obj)}")

result = json.dumps(Point(3, 4), default=point_encoder)
print(result)
"#;
    assert_eq!(run_python(src), vec![r#"{"x": 3, "y": 4}"#]);
}

#[test]
fn test_py_json_object_hook_decoder() {
    let src = r#"
import json

def as_point(d):
    if "x" in d and "y" in d:
        return (d["x"], d["y"])
    return d

result = json.loads('{"x": 3, "y": 4}', object_hook=as_point)
print(result)
print(type(result).__name__)
"#;
    assert_eq!(run_python(src), vec!["(3, 4)", "tuple"]);
}

#[test]
fn test_py_json_loads_nested_structures() {
    let src = r#"
import json

raw = '{"users": [{"name": "Alice", "roles": ["admin", "user"]}, {"name": "Bob", "roles": ["user"]}]}'
data = json.loads(raw)
print(data["users"][0]["name"])
print(data["users"][1]["roles"])
"#;
    assert_eq!(run_python(src), vec!["Alice", "['user']"]);
}

#[test]
fn test_py_json_dumps_special_floats() {
    let src = r#"
import json

try:
    json.dumps(float('inf'))
except ValueError as e:
    print("ValueError: Infinity not allowed")

result = json.dumps(float('inf'), allow_nan=True)
print(result)
"#;
    assert_eq!(
        run_python(src),
        vec!["ValueError: Infinity not allowed", "Infinity"]
    );
}

#[test]
fn test_py_json_roundtrip_list_and_nested_dict() {
    let src = r#"
import json

original = {"matrix": [[1, 2], [3, 4]], "meta": {"rows": 2, "cols": 2}}
rt = json.loads(json.dumps(original))
print(rt["matrix"][1][0])
print(rt["meta"]["rows"])
"#;
    assert_eq!(run_python(src), vec!["3", "2"]);
}

#[test]
fn test_py_json_object_pairs_hook() {
    let src = r#"
import json
from collections import OrderedDict

raw = '{"z": 1, "a": 2, "m": 3}'
ordered = json.loads(raw, object_pairs_hook=OrderedDict)
print(list(ordered.keys()))
"#;
    assert_eq!(run_python(src), vec!["['z', 'a', 'm']"]);
}

#[test]
fn test_py_json_decode_error_on_invalid() {
    let src = r#"
import json

try:
    json.loads("{'key': 'value'}")  # single quotes invalid in JSON
except json.JSONDecodeError as e:
    print("JSONDecodeError caught")
    print(e.pos >= 0)
"#;
    assert_eq!(run_python(src), vec!["JSONDecodeError caught", "True"]);
}

#[test]
fn test_py_json_encode_unicode() {
    let src = r#"
import json

data = {"greeting": "こんにちは", "emoji": "😀"}
j1 = json.dumps(data, ensure_ascii=True)
j2 = json.dumps(data, ensure_ascii=False)
parsed = json.loads(j1)
print(parsed["emoji"])
print("\\u" in j1)
print("\\u" in j2)
"#;
    assert_eq!(run_python(src), vec!["😀", "True", "False"]);
}

#[test]
fn test_py_json_streaming_with_stringio() {
    let src = r#"
import json, io

buf = io.StringIO()
json.dump([1, 2, 3], buf)
buf.seek(0)
result = json.load(buf)
print(result)
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]"]);
}
