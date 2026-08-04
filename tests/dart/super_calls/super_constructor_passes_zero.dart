// vybe-test: dart/super_calls/super_constructor_passes_zero
// origin: languages/dart/tests/dart/test_super_calls.rs

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

class Slot {
  int id;
  Slot(this.id);
}
class Empty extends Slot {
  Empty() : super(0);
}
void __vybeMain() {
  __p(Empty().id);
}

void main() {
  __vybeMain();
  __check('0');
}
