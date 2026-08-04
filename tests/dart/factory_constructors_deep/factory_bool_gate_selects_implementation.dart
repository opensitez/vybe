// vybe-test: dart/factory_constructors_deep/factory_bool_gate_selects_implementation
// origin: languages/dart/tests/dart/test_factory_constructors_deep.rs

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

class Result {
  int code;
  Result(this.code);
  factory Result.ok() {
    return Result(0);
  }
  factory Result.err() {
    return Result(1);
  }
}
void __vybeMain() {
  __p(Result.ok().code + Result.err().code);
}

void main() {
  __vybeMain();
  __check('1');
}
