// vybe-test: dart/typedefs_core/generic_typedef_nested_mapper_to_length
// origin: languages/dart/tests/dart/test_typedefs_core.rs

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

typedef ToLength = int Function(String);
typedef StringList = List<String>;
int totalChars(StringList words, ToLength measure) {
  var sum = 0;
  for (var word in words) {
    sum += measure(word);
  }
  return sum;
}
void __vybeMain() {
  __p(totalChars(['a', 'bb', 'ccc'], (s) => s.length));
}

void main() {
  __vybeMain();
  __check('6');
}
