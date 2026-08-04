// vybe-test: dart/flutter_widgets_list_view/list_view_builder
// origin: languages/dart/tests/dart/test_flutter_widgets_list_view.rs

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
  final lv = ListView.builder(
    itemCount: 10,
    itemBuilder: (BuildContext context, int index) => const SizedBox(),
  );
  __p(lv is BoxScrollView);
}

void main() {
  __vybeMain();
  __check('true');
}
