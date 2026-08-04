// vybe-test: dart/set_algebra_deep/intersection_of_three_overlapping_ranges
// origin: languages/dart/tests/dart/test_set_algebra_deep.rs

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
  var a = {for (var i = 1; i < 21; i++) i};
  var b = {for (var i = 5; i < 16; i++) i};
  var c = {for (var i = 8; i < 13; i++) i};
  var r = a.intersection(b).intersection(c).toList()..sort();
  __p(r.length);
  __p(r.join(','));
}

void main() {
  __vybeMain();
  __check('5\n8,9,10,11,12');
}
