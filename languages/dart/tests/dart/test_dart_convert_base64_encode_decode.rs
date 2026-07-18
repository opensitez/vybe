use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:convert Base64 encode/decode
// ═══════════════════════════════════════════════════════════

#[test]
fn base64_encode_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = [104, 101, 108, 108, 111]; // "hello"
  print(base64Encode(bytes));
}
"#
        ),
        vec!["aGVsbG8="]
    );
}

#[test]
fn base64_decode_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final decoded = base64Decode('aGVsbG8=');
  print(decoded.join('-'));
}
"#
        ),
        vec!["104-101-108-108-111"]
    );
}

#[test]
fn base64_encode_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  print(base64Encode([]));
}
"#
        ),
        vec![""]
    );
}

#[test]
fn base64_decode_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final decoded = base64Decode('');
  print(decoded.length);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn base64url_encode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Bytes that generate '+' and '/' in normal base64: [251, 239]
  // In base64url it should be '-' and '_'
  final bytes = [251, 239]; 
  // base64: ++8=
  // base64url: --8=
  print(base64UrlEncode(bytes));
}
"#
        ),
        vec!["--8="]
    );
}

#[test]
fn base64url_decode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final decoded = base64Url.decode('--8=');
  print('${decoded[0]}:${decoded[1]}');
}
"#
        ),
        vec!["251:239"]
    );
}

#[test]
fn base64_decode_invalid_length_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    base64Decode('aGVs'); // Valid length is multiple of 4, but let's test a badly padded string like 'a'
    base64Decode('a');
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}

#[test]
fn base64_decode_invalid_character_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    base64Decode('aGVsbG8$'); // $ is invalid
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}

#[test]
fn base64_decode_whitespace_ignored() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Dart's base64 ignores whitespace according to docs, actually wait - the standard base64 parser might throw on whitespace.
  // Wait, Dart 2.0+ base64Decode throws on whitespace. Let's verify exception.
  try {
    base64Decode('aGVs bG8=');
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}

#[test]
fn base64_encode_large_list() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = List<int>.filled(100, 65); // 'A'
  final encoded = base64Encode(bytes);
  print(encoded.length); // 100 bytes -> ceil(100/3)*4 = 34 * 4 = 136
}
"#
        ),
        vec!["136"]
    );
}

#[test]
fn base64_decode_padding_required() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    // Missing padding
    base64Decode('aGVsbG8');
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}

#[test]
fn base64_encode_all_byte_values() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = List<int>.generate(256, (i) => i);
  final encoded = base64Encode(bytes);
  final decoded = base64Decode(encoded);
  print(decoded.length == 256 && decoded[255] == 255);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn base64url_encode_no_padding() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // In Dart, you can strip padding manually, or sometimes url encode might not need it, 
  // but base64UrlEncode does add padding by default.
  final bytes = [104, 101, 108, 108, 111]; // "hello" -> aGVsbG8=
  print(base64UrlEncode(bytes).endsWith('='));
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn base64_decode_base64url_with_standard_decoder_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Decoding url-safe string with standard decoder
  try {
    base64.decode('--8=');
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}

#[test]
fn base64url_decode_standard_base64_with_url_decoder_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Decoding standard string with url-safe decoder
  try {
    base64Url.decode('++8=');
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}

#[test]
fn base64_decode_normalize() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Dart provides base64.normalize to add padding or fix string length
  final normalized = base64.normalize('aGVsbG8'); // missing '='
  print(normalized);
}
"#
        ),
        vec!["aGVsbG8="]
    );
}

#[test]
fn base64url_normalize() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final normalized = base64Url.normalize('--8');
  print(normalized);
}
"#
        ),
        vec!["--8="]
    );
}

#[test]
fn base64_normalize_invalid_length_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    base64.normalize('aGV'); // 3 chars, cannot be padded to 4 (requires at least 2 chars of data)
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}

#[test]
fn base64_codec_encode_decode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final codec = base64;
  final encoded = codec.encode([1, 2, 3]);
  final decoded = codec.decode(encoded);
  print(decoded[1]);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn base64_chunked_encode_decode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Testing converter
  final encoder = base64.encoder;
  final encoded = encoder.convert([10, 20]);
  print(encoded);
}
"#
        ),
        vec!["ChQ="]
    );
}
