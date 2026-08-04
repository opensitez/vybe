// vybe-test: dart/closures/closure_stored_in_list_and_invoked
// origin: languages/dart/tests/dart/test_closures.rs

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
  var fns = <int Function(int)>[
    (x) => x + 1,
    (x) => x * 2,
    (x) => x - 1,
  ];
  __p(fns[0](5));
  __p(fns[1](5));
  __p(fns[2](5));
}

void main() {
  __vybeMain();
  __check('6\n10\n4');
}
