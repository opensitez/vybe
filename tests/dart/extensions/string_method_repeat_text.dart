// vybe-test: dart/extensions/string_method_repeat_text
// origin: languages/dart/tests/dart/test_extensions.rs

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

extension StrRepeat on String {
  String repeatText(int n) {
    var out = '';
    for (var i = 0; i < n; i++) {
      out += this;
    }
    return out;
  }
}
void __vybeMain() {
  __p('a'.repeatText(2));
}

void main() {
  __vybeMain();
  __check('aa');
}
