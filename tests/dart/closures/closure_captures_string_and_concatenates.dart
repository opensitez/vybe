// vybe-test: dart/closures/closure_captures_string_and_concatenates
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
  var prefix = 'vybe';
  var label = (String suffix) => '$prefix-$suffix';
  __p(label('dart'));
}

void main() {
  __vybeMain();
  __check('vybe-dart');
}
