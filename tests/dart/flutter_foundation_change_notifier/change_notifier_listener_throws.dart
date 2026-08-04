// vybe-test: dart/flutter_foundation_change_notifier/change_notifier_listener_throws
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
void __vybeMain() {
  final notifier = ChangeNotifier();
  int count = 0;
  notifier.addListener(() { throw Exception('Bad listener'); });
  notifier.addListener(() { count++; });
  try {
    notifier.notifyListeners();
  } catch (e) {
    // Depending on Flutter implementation, it might catch and report to FlutterError,
    // or bubble up if there is no zone handler.
    __p('Exception handled');
  }
  __p(count);
}

void main() {
  __vybeMain();
  __check('Exception handled\n1');
}
