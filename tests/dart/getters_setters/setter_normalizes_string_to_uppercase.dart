// vybe-test: dart/getters_setters/setter_normalizes_string_to_uppercase
// origin: languages/dart/tests/dart/test_getters_setters.rs

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
  String _value = '';
  String get value {
    return _value;
  }
  set value(String v) {
    _value = v.toUpperCase();
  }
}
void __vybeMain() {
  var c = Code();
  c.value = 'abc';
  __p(c.value);
}

void main() {
  __vybeMain();
  __check('ABC');
}
