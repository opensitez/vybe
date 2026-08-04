// vybe-test: dart/flutter_ui_paragraph_builder/paragraph_style_properties
// origin: languages/dart/tests/dart/test_flutter_ui_paragraph_builder.rs

final StringBuffer __vybeOut = StringBuffer();

void __p(Object? o) {
  __vybeOut.writeln(o);
}

void __check(String want) {
  var got = __vybeOut.toString();
  // `writeln` on the final print contributes a trailing newline that the
  // expected line vector never carried.
  if (got.endsWith('\n')) {
    got = got.substring(0, got.length - 1);
  }
  if (got != want) {
    print('FAIL: want [$want] got [$got]');
    throw Exception('assertion failed');
  }
}

import 'dart:ui';
void __vybeMain() {
  final style = ParagraphStyle(
    textAlign: TextAlign.center,
    textDirection: TextDirection.rtl,
    maxLines: 2,
    fontFamily: 'Roboto',
    fontSize: 16.0,
    fontWeight: FontWeight.bold,
    fontStyle: FontStyle.italic,
  );
  __p(style != null);
}

void main() {
  __vybeMain();
  __check('true');
}
