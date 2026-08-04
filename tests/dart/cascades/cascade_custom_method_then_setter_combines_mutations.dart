// vybe-test: dart/cascades/cascade_custom_method_then_setter_combines_mutations
// origin: languages/dart/tests/dart/test_cascades.rs

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

class Widget {
  String label = '';
  void setLabel(String s) { label = s; }
  set prefix(String p) { label = p + label; }
}
void __vybeMain() {
  var w = Widget();
  w..setLabel('base')..prefix = 'pre-';
  __p(w.label);
}

void main() {
  __vybeMain();
  __check('pre-base');
}
