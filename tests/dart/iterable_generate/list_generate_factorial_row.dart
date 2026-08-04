// vybe-test: dart/iterable_generate/list_generate_factorial_row
// origin: languages/dart/tests/dart/test_iterable_generate.rs

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
  var list = List.generate(4, (i) {
    var f = 1;
    for (var j = 1; j <= i; j++) {
      f = f * j;
    }
    return f;
  });
  __p(list.join(','));
}

void main() {
  __vybeMain();
  __check('1,1,2,6');
}
