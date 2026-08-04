// vybe-test: dart/flutter_widgets_inherited_widget/inherited_model_subclass
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
class MyModel extends InheritedModel<String> {
  MyModel({required Widget child}) : super(child: child);
  @override
  bool updateShouldNotify(MyModel oldWidget) => true;
  @override
  bool updateShouldNotifyDependent(MyModel oldWidget, Set<String> dependencies) {
    return dependencies.contains('update');
  }
}
void __vybeMain() {
  final w = MyModel(child: const SizedBox());
  __p(w != null);
}

void main() {
  __vybeMain();
  __check('true');
}
