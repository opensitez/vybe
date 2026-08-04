// vybe-test: dart/flutter_ui_paragraph_builder/paragraph_longest_line
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
  final builder = ParagraphBuilder(ParagraphStyle());
  builder.addText('A very long line indeed');
  final paragraph = builder.build();
  paragraph.layout(ParagraphConstraints(width: 1000.0));
  __p(paragraph.longestLine >= 0.0);
}

void main() {
  __vybeMain();
  __check('true');
}
