use super::helpers::run_python;

#[test]
fn test_python_json_default_encoder() {
    let src = r#"
import json
from datetime import datetime

def encoder(obj):
    if isinstance(obj, datetime):
        return obj.strftime('%Y-%m-%d')
    raise TypeError

print(json.dumps({'day': datetime(2026, 7, 26)}, default=encoder))
"#;
    assert_eq!(run_python(src), vec!["{\"day\": \"2026-07-26\"}"]);
}

#[test]
fn test_python_json_custom_class() {
    let src = r#"
import json
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

def encode(o):
    if isinstance(o, Point):
        return {'x': o.x, 'y': o.y}
    raise TypeError

print(json.dumps(Point(1, 2), default=encode))
"#;
    assert_eq!(run_python(src), vec!["{\"x\": 1, \"y\": 2}"]);
}
