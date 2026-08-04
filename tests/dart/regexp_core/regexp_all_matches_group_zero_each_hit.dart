// vybe-test: dart/regexp_core/regexp_all_matches_group_zero_each_hit
// origin: languages/dart/tests/dart/test_regexp_core.rs

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
  var re = RegExp(r'\d+');
  var sum = 0;
  for (var m in re.allMatches('a1b22')) {
    sum += int.parse(m.group(0)!);
  }
  __p(sum);
}

void main() {
  __vybeMain();
  __check('23');
}
