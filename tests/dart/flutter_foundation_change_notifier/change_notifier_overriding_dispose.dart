// vybe-test: dart/flutter_foundation_change_notifier/change_notifier_overriding_dispose
// origin: languages/dart/tests/dart/test_flutter_foundation_change_notifier.rs

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

import 'package:flutter/foundation.dart';
class BadNotifier extends ChangeNotifier {
  @override
  void dispose() {
    // missing super.dispose()
  }
}
void __vybeMain() {
  final n = BadNotifier();
  n.dispose();
  // Because super was not called, it doesn't throw on notify
  n.notifyListeners();
  print('survived');
}

void main() {
  __vybeMain();
  __check('survived');
}
