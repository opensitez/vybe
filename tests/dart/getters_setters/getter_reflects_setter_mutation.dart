// vybe-test: dart/getters_setters/getter_reflects_setter_mutation
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

class Label {
  String _text = 'a';
  String get text {
    return _text;
  }
  set text(String v) {
    _text = v;
  }
}
void __vybeMain() {
  var l = Label();
  l.text = 'updated';
  __p(l.text);
}

void main() {
  __vybeMain();
  __check('updated');
}
