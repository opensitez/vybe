// vybe-test: dart/typedefs_core/generic_typedef_callback_list_invocation_order
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

typedef Consumer<T> = void Function(T);
void __vybeMain() {
  var log = <int>[];
  List<Consumer<int>> steps = [
    (n) => log.add(1),
    (n) => log.add(2),
    (n) => log.add(3),
  ];
  for (var step in steps) {
    step(0);
  }
  __p(log.join(','));
}

void main() {
  __vybeMain();
  __check('1,2,3');
}
