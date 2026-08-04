// vybe-test: dart/generics_core/generic_result_err_variant
// origin: languages/dart/tests/dart/test_generics_core.rs

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

class Result<T, E> {
  T? value;
  E? error;
  Result.ok(this.value) : error = null;
  Result.err(this.error) : value = null;
  bool get isOk {
    return error == null;
  }
}
void __vybeMain() {
  var r = Result<int, String>.err('fail');
  __p(r.isOk);
  __p(r.error);
}

void main() {
  __vybeMain();
  __check('false\nfail');
}
