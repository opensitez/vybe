// vybe-test: dart/operator_overloading/operator_plus_string_concat_fields
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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

class Tag {
  String a;
  String b;
  Tag(this.a, this.b);
  Tag operator +(Tag other) {
    return Tag(a + other.a, b + other.b);
  }
}
void __vybeMain() {
  var t = Tag('x', '1') + Tag('y', '2');
  __p(t.a + t.b);
}

void main() {
  __vybeMain();
  __check('xy12');
}
