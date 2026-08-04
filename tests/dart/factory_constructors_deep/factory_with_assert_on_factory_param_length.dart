// vybe-test: dart/factory_constructors_deep/factory_with_assert_on_factory_param_length
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class Code {
  String text;
  Code._(this.text);
  factory Code(String t) {
    assert(t.length >= 2);
    return Code._(t);
  }
}
void __vybeMain() {
  __p(Code('xy').text.length);
}

void main() {
  __vybeMain();
  __check('2');
}
