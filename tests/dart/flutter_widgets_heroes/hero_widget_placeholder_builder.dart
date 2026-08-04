// vybe-test: dart/flutter_widgets_heroes/hero_widget_placeholder_builder
// origin: languages/dart/tests/dart/test_flutter_widgets_heroes.rs

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
  Widget builder(BuildContext context, Size heroSize, Widget child) {
    return const SizedBox();
  }
  final h = Hero(tag: 'tag', child: const SizedBox(), placeholderBuilder: builder);
  __p(h.placeholderBuilder != null);
}

void main() {
  __vybeMain();
  __check('true');
}
