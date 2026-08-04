// vybe-test: dart/top_level_members/top_level_pipeline_three_calls
// origin: languages/dart/tests/dart/test_top_level_members.rs

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

int parse(String s) {
  return int.parse(s);
}
int doubleIt(int n) {
  return n * 2;
}
int finish(int n) {
  return n + 1;
}
void __vybeMain() {
  __p(finish(doubleIt(parse('10'))));
}

void main() {
  __vybeMain();
  __check('21');
}
