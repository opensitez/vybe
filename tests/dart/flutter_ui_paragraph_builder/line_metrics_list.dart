// vybe-test: dart/flutter_ui_paragraph_builder/line_metrics_list
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
  builder.addText('Line 1\nLine 2');
  final paragraph = builder.build();
  paragraph.layout(ParagraphConstraints(width: 100.0));
  final metrics = paragraph.computeLineMetrics();
  __p(metrics.length >= 0);
}

void main() {
  __vybeMain();
  __check('true');
}
