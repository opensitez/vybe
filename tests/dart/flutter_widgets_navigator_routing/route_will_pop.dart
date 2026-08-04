// vybe-test: dart/flutter_widgets_navigator_routing/route_will_pop
// origin: languages/dart/tests/dart/test_flutter_widgets_navigator_routing.rs

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
void __vybeMain() async {
  final r = PageRouteBuilder<int>(pageBuilder: (c, a, sa) => const SizedBox());
  final res = await r.willPop();
  __p(res == RoutePopDisposition.pop);
}

Future<void> main() async {
  await __vybeMain();
  __check('true');
}
