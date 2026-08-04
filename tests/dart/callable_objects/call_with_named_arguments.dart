// vybe-test: dart/callable_objects/call_with_named_arguments
// origin: languages/dart/tests/dart/test_callable_objects.rs

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

class Config {
  String call({required String mode, required int level}) {
    return '$mode:$level';
  }
}
void __vybeMain() {
  __p(Config()(mode: 'fast', level: 3));
}

void main() {
  __vybeMain();
  __check('fast:3');
}
