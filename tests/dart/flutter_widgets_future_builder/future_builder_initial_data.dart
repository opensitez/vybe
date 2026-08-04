// vybe-test: dart/flutter_widgets_future_builder/future_builder_initial_data
// origin: languages/dart/tests/dart/test_flutter_widgets_future_builder.rs

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
  final fb = FutureBuilder<int>(
    initialData: 42,
    future: Future.value(100),
    builder: (BuildContext context, AsyncSnapshot<int> snapshot) {
      return const SizedBox();
    },
  );
  __p(fb.initialData);
}

void main() {
  __vybeMain();
  __check('42');
}
