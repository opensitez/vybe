// vybe-test: dart/named_parameters/named_params_in_method_on_class
// origin: languages/dart/tests/dart/test_named_parameters.rs

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

class Printer {
  void show({String text = 'none'}) {
    __p(text);
  }
}
void __vybeMain() {
  Printer().show(text: 'line');
}

void main() {
  __vybeMain();
  __check('line');
}
