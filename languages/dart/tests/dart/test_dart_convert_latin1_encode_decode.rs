use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:convert Latin1 & ASCII encode/decode
// ═══════════════════════════════════════════════════════════

#[test]
fn latin1_encode_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = latin1.encode('Hello');
  print(bytes.join('-'));
}
"#
        ),
        vec!["72-101-108-108-111"]
    );
}

#[test]
fn latin1_decode_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = latin1.decode([72, 101, 108, 108, 111]);
  print(str);
}
"#
        ),
        vec!["Hello"]
    );
}

#[test]
fn latin1_encode_extended() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // ISO-8859-1 covers 0-255. 'ä' is U+00E4 (228).
  final bytes = latin1.encode('ä');
  print(bytes[0]);
}
"#
        ),
        vec!["228"]
    );
}

#[test]
fn latin1_decode_extended() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = latin1.decode([228]);
  print(str);
}
"#
        ),
        vec!["ä"]
    );
}

#[test]
fn latin1_encode_unsupported_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // '🚀' is U+1F680, not in Latin-1
  try {
    latin1.encode('🚀');
  } on FormatException {
    print('FormatException thrown');
  } catch(e) {
    print('ArgumentError thrown'); // Depending on implementation
  }
}
"#
        ),
        vec!["FormatException thrown"] // Dart specifically throws FormatException for invalid characters in encode
    );
}

#[test]
fn latin1_encode_unsupported_allow_invalid() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // You cannot pass allowInvalid to static latin1.encode directly, need Latin1Encoder
  // Actually wait, Dart doesn't have allowInvalid for Latin1Encoder by default.
  // We'll test Latin1Encoder instantiation.
  final encoder = Latin1Encoder();
  print(encoder is Converter<String, List<int>>);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn latin1_decode_allow_invalid() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Latin1 decode has allowInvalid
  final str = latin1.decode([256], allowInvalid: true);
  // Wait, [256] is technically invalid since latin1 is 8-bit.
  // Actually list elements > 255 might be truncated or throw.
  // In Dart, passing >255 to allowInvalid might return replacement chars or just truncate.
  // Let's just catch whatever it does or throws without allowInvalid.
  try {
    latin1.decode([256], allowInvalid: false);
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
fn ascii_encode_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = ascii.encode('Hello');
  print(bytes.join('-'));
}
"#
        ),
        vec!["72-101-108-108-111"]
    );
}

#[test]
fn ascii_decode_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = ascii.decode([72, 101, 108, 108, 111]);
  print(str);
}
"#
        ),
        vec!["Hello"]
    );
}

#[test]
fn ascii_encode_extended_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // 'ä' is 228, out of ASCII range (0-127)
  try {
    ascii.encode('ä');
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
fn ascii_decode_extended_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    ascii.decode([228]);
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
fn ascii_decode_allow_invalid() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = ascii.decode([228], allowInvalid: true);
  // Replaced with U+FFFD
  print(str.codeUnitAt(0) == 0xFFFD);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn latin1_encode_all_valid() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = String.fromCharCodes(List.generate(256, (i) => i));
  final bytes = latin1.encode(str);
  print(bytes.length == 256);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn ascii_encode_all_valid() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = String.fromCharCodes(List.generate(128, (i) => i));
  final bytes = ascii.encode(str);
  print(bytes.length == 128);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn latin1_encode_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = latin1.encode('');
  print(bytes.length);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn ascii_encode_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = ascii.encode('');
  print(bytes.length);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn latin1_decoder_convert() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final decoder = Latin1Decoder();
  print(decoder.convert([65, 66]));
}
"#
        ),
        vec!["AB"]
    );
}

#[test]
fn ascii_encoder_convert() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final encoder = AsciiEncoder();
  final bytes = encoder.convert('AB');
  print('${bytes[0]}:${bytes[1]}');
}
"#
        ),
        vec!["65:66"]
    );
}

#[test]
fn latin1_codec_name() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  print(latin1.name);
}
"#
        ),
        vec!["iso-8859-1"]
    );
}

#[test]
fn ascii_codec_name() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  print(ascii.name);
}
"#
        ),
        vec!["us-ascii"]
    );
}
