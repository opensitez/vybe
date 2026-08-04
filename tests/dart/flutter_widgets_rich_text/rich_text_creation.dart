// vybe-test: dart/flutter_widgets_rich_text/rich_text_creation
// origin: languages/dart/tests/dart/test_flutter_widgets_rich_text.rs

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

import 'package:flutter/widgets.dart';
void __vybeMain() {
  final rt = RichText(text: const TextSpan(text: 'Hello'));
  __p(rt.text.toPlainText());
}

void main() {
  __vybeMain();
  __check('Hello');
}
