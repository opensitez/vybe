// vybe-test: dart/class_modifiers/sealed_subtype_with_method
// origin: languages/dart/tests/dart/test_class_modifiers.rs

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

sealed class Msg {}
class TextMsg extends Msg {
  String text;
  TextMsg(this.text);
  int length() {
    return text.length;
  }
}
void __vybeMain() {
  __p(TextMsg('hi').length());
}

void main() {
  __vybeMain();
  __check('2');
}
