// vybe-test: dart/super_calls/super_method_concat_three_parts
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

class Part {
  String mid() {
    return 'b';
  }
}
class Full extends Part {
  String mid() {
    return 'a' + super.mid() + 'c';
  }
}
void __vybeMain() {
  __p(Full().mid());
}

void main() {
  __vybeMain();
  __check('abc');
}
