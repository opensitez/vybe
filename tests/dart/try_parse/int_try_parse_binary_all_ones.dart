// vybe-test: dart/try_parse/int_try_parse_binary_all_ones
// origin: languages/dart/tests/dart/test_try_parse.rs

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

void __vybeMain() {
  __p(int.tryParse('11111111', radix: 2));
}

void main() {
  __vybeMain();
  __check('255');
}
