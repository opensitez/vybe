// vybe-test: dart/factory_constructors_deep/factory_parses_enum_like_string
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

class Mode {
  String value;
  Mode._(this.value);
  factory Mode.parse(String s) {
    if (s == 'on') {
      return Mode._('on');
    }
    return Mode._('off');
  }
}
void __vybeMain() {
  __p(Mode.parse('on').value);
}

void main() {
  __vybeMain();
  __check('on');
}
