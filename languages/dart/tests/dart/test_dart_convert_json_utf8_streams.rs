use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Dart: dart:convert JSON & UTF8 Streams
// ═══════════════════════════════════════════════════════════

#[test]
fn json_utf8_encode_stream() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable([{'a': 1}, {'b': 2}]);
  final out = await stream.transform(json.encoder).transform(utf8.encoder).toList();
  // out is List<List<int>>
  print(out.length);
}
"#
        ),
        vec!["2"]
    );
}

#[test]
fn json_utf8_decode_stream() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable([utf8.encode('{"a":1}')]);
  final out = await stream.transform(utf8.decoder).transform(json.decoder).toList();
  print(out[0]['a']);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn utf8_decoder_stream_chunked() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  // 'ä' is 0xC3, 0xA4
  final stream = Stream.fromIterable([ [0xC3], [0xA4] ]);
  final out = await stream.transform(utf8.decoder).join();
  print(out);
}
"#
        ),
        vec!["ä"]
    );
}

#[test]
fn utf8_encoder_stream() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable(['h', 'i']);
  final out = await stream.transform(utf8.encoder).toList();
  print('${out[0][0]}:${out[1][0]}'); // 104, 105
}
"#
        ),
        vec!["104:105"]
    );
}

#[test]
fn json_utf8_bind() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable([ [123, 34, 97, 34, 58, 49, 125] ]); // {"a":1}
  final decoded = await utf8.decoder.bind(stream).transform(json.decoder).first;
  print(decoded['a']);
}
"#
        ),
        vec!["1"]
    );
}

#[test]
fn utf8_decode_malformed_stream_allow_malformed() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final decoder = Utf8Decoder(allowMalformed: true);
  final stream = Stream.fromIterable([[0xFF]]); // Invalid UTF-8
  final out = await stream.transform(decoder).join();
  print(out.length > 0); // Contains replacement character
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn utf8_decode_malformed_stream_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final decoder = Utf8Decoder(allowMalformed: false);
  final stream = Stream.fromIterable([[0xFF]]); // Invalid UTF-8
  try {
    await stream.transform(decoder).join();
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
fn json_decode_stream_multiple_objects_fails() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable(['{"a":1}{"b":2}']);
  try {
    await stream.transform(json.decoder).toList();
  } on FormatException {
    print('FormatException thrown');
  }
}
"#
        ),
        vec!["FormatException thrown"] // JsonDecoder expects exactly one JSON value per chunk/stream unless it's a specific parser
    );
}

#[test]
fn line_splitter_stream() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable(['line1\nli', 'ne2\r\nline3']);
  final lines = await stream.transform(const LineSplitter()).toList();
  print(lines.length);
  print(lines[1]);
}
"#
        ),
        vec!["3\nline2"]
    );
}

#[test]
fn line_splitter_bind() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable(['A\nB']);
  final lines = await const LineSplitter().bind(stream).toList();
  print(lines.join('-'));
}
"#
        ),
        vec!["A-B"]
    );
}

#[test]
fn utf8_decode_stream_empty() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream<List<int>>.empty();
  final out = await stream.transform(utf8.decoder).join();
  print(out.isEmpty);
}
"#
        ),
        vec!["true"]
    );
}

#[test]
fn json_decode_stream_empty_throws() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream<String>.empty();
  try {
    await stream.transform(json.decoder).first;
  } catch(e) {
    print('StateError thrown'); // first on empty stream throws StateError
  }
}
"#
        ),
        vec!["StateError thrown"]
    );
}

#[test]
fn chunked_conversion_sink_json() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  var results = [];
  var sink = json.encoder.startChunkedConversion(
    ChunkedConversionSink.withCallback((chunks) {
      results.addAll(chunks);
    })
  );
  sink.add({'a': 1});
  sink.close();
  print(results[0]);
}
"#
        ),
        vec![r#"{"a":1}"#]
    );
}

#[test]
fn chunked_conversion_sink_utf8() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  var results = [];
  var sink = utf8.encoder.startChunkedConversion(
    ByteConversionSink.withCallback((bytes) {
      results.addAll(bytes);
    })
  );
  sink.add('A');
  sink.close();
  print(results[0]);
}
"#
        ),
        vec!["65"]
    );
}

#[test]
fn string_conversion_sink_utf8_decode() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  var result = '';
  var sink = utf8.decoder.startChunkedConversion(
    StringConversionSink.withCallback((str) {
      result += str;
    })
  );
  sink.add([65, 66]);
  sink.close();
  print(result);
}
"#
        ),
        vec!["AB"]
    );
}

#[test]
fn string_conversion_sink_as_utf8_sink() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  var result = '';
  var outSink = StringConversionSink.withCallback((s) => result += s);
  var inSink = outSink.asUtf8Sink(false);
  inSink.add([67, 68]);
  inSink.close();
  print(result);
}
"#
        ),
        vec!["CD"]
    );
}

#[test]
fn string_conversion_sink_from_string_sink() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() {
  var buffer = StringBuffer();
  var sink = StringConversionSink.fromStringSink(buffer);
  sink.add('X');
  sink.add('Y');
  sink.close();
  print(buffer.toString());
}
"#
        ),
        vec!["XY"]
    );
}

#[test]
fn base64_encoder_stream() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable([ [1, 2, 3] ]);
  final out = await stream.transform(base64.encoder).join();
  print(out);
}
"#
        ),
        vec!["AQID"]
    );
}

#[test]
fn base64_decoder_stream() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable(['AQID']);
  final out = await stream.transform(base64.decoder).toList();
  print('${out[0][0]}:${out[0][1]}');
}
"#
        ),
        vec!["1:2"]
    );
}

#[test]
fn ascii_encoder_stream() {
    assert_eq!(
        run_prints(
            r#"
import 'dart:convert';
void main() async {
  final stream = Stream.fromIterable(['A', 'B']);
  final out = await stream.transform(ascii.encoder).toList();
  print('${out[0][0]}:${out[1][0]}');
}
"#
        ),
        vec!["65:66"]
    );
}
