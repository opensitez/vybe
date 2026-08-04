// vybe-test: dart/null_operators/coalesce_nested_with_null_aware_access
// origin: languages/dart/tests/dart/test_null_operators.rs

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

class Node { String? name; Node? next; }
void __vybeMain() {
  var n = Node();
  __p(n.next?.name ?? n.name ?? 'anon');
}

void main() {
  __vybeMain();
  __check('anon');
}
