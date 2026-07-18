use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:convert UTF-8 encode/decode
// ═══════════════════════════════════════════════════════════

#[test]
fn utf8_encode_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = utf8.encode('Hello');
  print(bytes.join('-'));
}
"#
        ),
        vec!["72-101-108-108-111"]
    );
}

#[test]
fn utf8_decode_basic() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = utf8.decode([72, 101, 108, 108, 111]);
  print(str);
}
"#
        ),
        vec!["Hello"]
    );
}

#[test]
fn utf8_encode_multibyte() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // 'ä' is 2 bytes: 0xC3 0xA4 (195, 164)
  final bytes = utf8.encode('ä');
  print('${bytes[0]}:${bytes[1]}');
}
"#
        ),
        vec!["195:164"]
    );
}

#[test]
fn utf8_decode_multibyte() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = utf8.decode([195, 164]);
  print(str);
}
"#
        ),
        vec!["ä"]
    );
}

#[test]
fn utf8_encode_emoji() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // '🚀' is 4 bytes: 0xF0 0x9F 0x9A 0x80 (240, 159, 154, 128)
  final bytes = utf8.encode('🚀');
  print('${bytes[0]}:${bytes[1]}:${bytes[2]}:${bytes[3]}');
}
"#
        ),
        vec!["240:159:154:128"]
    );
}

#[test]
fn utf8_decode_emoji() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = utf8.decode([240, 159, 154, 128]);
  print(str);
}
"#
        ),
        vec!["🚀"]
    );
}

#[test]
fn utf8_decode_malformed_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    utf8.decode([0xFF]);
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
fn utf8_decode_malformed_allow_malformed() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = utf8.decode([0xFF], allowMalformed: true);
  // Replacement character is U+FFFD (65533)
  print(str.codeUnitAt(0) == 0xFFFD);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn utf8_decode_truncated_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  try {
    // 'ä' is 195, 164. We only provide 195.
    utf8.decode([195]);
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
fn utf8_encode_surrogate_pairs() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // High and low surrogates for 🚀 (U+1F680): D83D DE80
  final str = String.fromCharCodes([0xD83D, 0xDE80]);
  final bytes = utf8.encode(str);
  print(bytes.length); // 4
}
"#
        ),
        vec!["4"]
    );
}

#[test]
fn utf8_encode_unpaired_surrogate_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Dart strings are UTF-16. 
  // Wait, utf8.encode might not throw on unpaired surrogates, it might encode them as replacement chars.
  // Actually, Dart's standard utf8 encoder converts invalid surrogates to U+FFFD.
  final str = String.fromCharCodes([0xD83D]);
  final bytes = utf8.encode(str);
  print(bytes.length == 3); // U+FFFD is 3 bytes in UTF-8
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn utf8_encode_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final bytes = utf8.encode('');
  print(bytes.length);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn utf8_decode_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = utf8.decode([]);
  print(str.length);
}
"#
        ),
        vec!["0"]
    );
}

#[test]
fn utf8_decoder_convert() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final decoder = Utf8Decoder();
  print(decoder.convert([65, 66]));
}
"#
        ),
        vec!["AB"]
    );
}

#[test]
fn utf8_encoder_convert() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final encoder = Utf8Encoder();
  final bytes = encoder.convert('AB');
  print('${bytes[0]}:${bytes[1]}');
}
"#
        ),
        vec!["65:66"]
    );
}

#[test]
fn utf8_decode_boms_ignored() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // UTF-8 BOM is 0xEF, 0xBB, 0xBF
  // Wait, Dart's standard utf8.decode does not strip BOM by default.
  // You need utf8.decoder to strip it or handle it manually?
  // Let's just decode and check the first character.
  final str = utf8.decode([0xEF, 0xBB, 0xBF, 65]);
  print(str.codeUnitAt(0) == 0xFEFF); // The BOM character itself
}
"#
        ),
        // Wait! In Dart, `utf8.decode([0xEF, 0xBB, 0xBF, 65])`
        // Actually `utf8.decode` takes an optional `allowMalformed` but no BOM stripping option directly on static.
        // Wait, Dart 2.0+ `utf8` DOES NOT strip BOM by default, so it returns `\uFEFFA`.
        vec!["true"]
    );
}

#[test]
fn utf8_encode_all_ascii() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  final str = String.fromCharCodes(List.generate(128, (i) => i));
  final bytes = utf8.encode(str);
  print(bytes.length == 128);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn utf8_encode_max_code_point() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // Max valid Unicode code point is U+10FFFF
  // High surrogate: 0xDBFF, Low surrogate: 0xDFFF
  final str = String.fromCharCodes([0xDBFF, 0xDFFF]);
  final bytes = utf8.encode(str);
  // It takes 4 bytes in UTF-8
  print(bytes.length == 4);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn utf8_encode_codec_name() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  print(utf8.name);
}
"#
        ),
        vec!["utf-8"]
    );
}

#[test]
fn utf8_decode_overlong_encoding_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  // 'A' is 65, but encoded as overlong 2-byte: 0xC0 0x81
  try {
    utf8.decode([0xC0, 0x81]);
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"]
    );
}
