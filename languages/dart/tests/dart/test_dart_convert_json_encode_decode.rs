use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:convert JSON encode/decode
// ═══════════════════════════════════════════════════════════

#[test]
fn json_encode_basic_map() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final map = {'name': 'Dart', 'year': 2011};
  print(jsonEncode(map));
}
"#
        ),
        vec![r#"{"name":"Dart","year":2011}"#]
    );
}

#[test]
fn json_decode_basic_map() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final String jsonStr = '{"name":"Dart","year":2011}';
  final map = jsonDecode(jsonStr);
  print('${map['name']}:${map['year']}');
}
"#
        ),
        vec!["Dart:2011"]
    );
}

#[test]
fn json_encode_list() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final list = [1, 'two', false, null];
  print(jsonEncode(list));
}
"#
        ),
        vec![r#"[1,"two",false,null]"#]
    );
}

#[test]
fn json_decode_list() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final list = jsonDecode('[1,"two",false,null]');
  print('${list[0]}:${list[1]}:${list[2]}:${list[3]}');
}
"#
        ),
        vec!["1:two:false:null"]
    );
}

#[test]
fn json_encode_nested() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final data = {
    'users': [
      {'id': 1, 'active': true},
      {'id': 2, 'active': false}
    ]
  };
  print(jsonEncode(data));
}
"#
        ),
        vec![r#"{"users":[{"id":1,"active":true},{"id":2,"active":false}]}"#]
    );
}

#[test]
fn json_decode_nested() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = '{"users":[{"id":1,"active":true},{"id":2,"active":false}]}';
  final data = jsonDecode(str);
  print(data['users'][1]['id']);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn json_encode_with_to_json() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
class User {
  final String name;
  User(this.name);
  Map<String, dynamic> toJson() => {'name': name};
}
void main() {
  final user = User('Bob');
  print(jsonEncode(user));
}
"#
        ),
        vec![r#"{"name":"Bob"}"#]
    );
}

#[test]
fn json_encode_unsupported_object_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
class BadObject {}
void main() {
  try {
    jsonEncode(BadObject());
  } on JsonUnsupportedObjectError {
    print('JsonUnsupportedObjectError thrown');
  }
}
"#
        ),
        vec!["JsonUnsupportedObjectError thrown"]
    );
}

#[test]
fn json_encode_custom_encoder() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
class BadObject {}
void main() {
  final obj = BadObject();
  final result = jsonEncode(obj, toEncodable: (o) => 'fallback');
  print(result);
}
"#
        ),
        vec![r#""fallback""#]
    );
}

#[test]
fn json_decode_invalid_json_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    jsonDecode('{bad_json: 1}');
  } on FormatException catch (e) {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}

#[test]
fn json_decode_with_reviver() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final jsonStr = '{"a": 1, "b": 2}';
  final result = jsonDecode(jsonStr, reviver: (key, value) {
    if (key == 'a') return value * 10;
    return value;
  });
  print(result['a']);
}
"#
        ),
        vec!["10"]
    );
}

#[test]
fn json_encode_unicode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final map = {'icon': '🚀'};
  print(jsonEncode(map));
}
"#
        ),
        vec![r#"{"icon":"🚀"}"#]
    );
}

#[test]
fn json_encode_escape_characters() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final map = {'text': 'line1\nline2\t"quoted"'};
  print(jsonEncode(map));
}
"#
        ),
        vec![r#"{"text":"line1\nline2\t\"quoted\""}"#]
    );
}

#[test]
fn json_encode_double_nan_infinity() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    jsonEncode([double.nan, double.infinity]);
  } on JsonUnsupportedObjectError {
    print('JsonUnsupportedObjectError thrown');
  }
}
"#
        ),
        vec!["JsonUnsupportedObjectError thrown"]
    );
}

#[test]
fn json_encoder_indent() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final encoder = JsonEncoder.withIndent('  ');
  final map = {'a': 1};
  print(encoder.convert(map));
}
"#
        ),
        vec!["{\n  \"a\": 1\n}"]
    );
}

#[test]
fn json_decoder_convert() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final decoder = JsonDecoder();
  final list = decoder.convert('[1, 2]');
  print(list.length);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn json_decode_numbers() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final list = jsonDecode('[42, 3.14, -1]');
  print('${list[0] is int}:${list[1] is double}');
}
"#
        ),
        vec!["true:true"]
    );
}

#[test]
fn json_encode_large_integer() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Dart ints are 64-bit, JS JSON limits safely to 53-bit. 
  // Dart's jsonEncode on native supports full 64-bit int serialization
  final map = {'big': 9007199254740992}; // 2^53
  print(jsonEncode(map));
}
"#
        ),
        vec![r#"{"big":9007199254740992}"#]
    );
}

#[test]
fn json_codec_encode_decode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final encoded = json.encode([1, 2, 3]);
  final decoded = json.decode(encoded);
  print(decoded[1]);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn json_encode_cyclic_reference_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final list = [];
  list.add(list); // cyclic
  try {
    jsonEncode(list);
    // Usually crashes or throws
  } catch(e) {
    print('JsonUnsupportedObjectError thrown');
  }
}
"#
        ),
        vec!["JsonUnsupportedObjectError thrown"]
    );
}
