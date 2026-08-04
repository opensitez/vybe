// vybe-test: dart/object_protocol/custom_hashcode_overrides_default
// origin: languages/dart/tests/dart/test_object_protocol.rs

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

class Key {
  int k;
  Key(this.k);
  int get hashCode => k;
}
void __vybeMain() {
  __p(Key(3).hashCode);
}

void main() {
  __vybeMain();
  __check('3');
}
