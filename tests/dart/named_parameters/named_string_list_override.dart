// vybe-test: dart/named_parameters/named_string_list_override
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

void tags({List<String> items = const ['a']}) {
  __p(items.join(','));
}
void __vybeMain() {
  tags(items: ['x', 'y']);
}

void main() {
  __vybeMain();
  __check('x,y');
}
