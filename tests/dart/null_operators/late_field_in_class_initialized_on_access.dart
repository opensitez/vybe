// vybe-test: dart/null_operators/late_field_in_class_initialized_on_access
// origin: languages/dart/tests/dart/test_null_operators.rs

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

class Holder {
  late String tag;
  Holder() { tag = 'ready'; }
}
void __vybeMain() {
  __p(Holder().tag);
}

void main() {
  __vybeMain();
  __check('ready');
}
