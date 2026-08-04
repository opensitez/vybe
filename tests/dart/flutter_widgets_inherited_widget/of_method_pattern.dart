// vybe-test: dart/flutter_widgets_inherited_widget/of_method_pattern
// origin: languages/dart/tests/dart/test_flutter_widgets_inherited_widget.rs

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
class MyInherited extends InheritedWidget {
  final int value = 42;
  MyInherited({required Widget child}) : super(child: child);
  @override
  bool updateShouldNotify(MyInherited old) => false;
  static MyInherited? of(BuildContext context) {
    return context.dependOnInheritedWidgetOfExactType<MyInherited>();
  }
}
void __vybeMain() {
  __p('compiles');
}

void main() {
  __vybeMain();
  __check('compiles');
}
