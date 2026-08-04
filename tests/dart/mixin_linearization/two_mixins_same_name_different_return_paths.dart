// vybe-test: dart/mixin_linearization/two_mixins_same_name_different_return_paths
// origin: languages/dart/tests/dart/test_mixin_linearization.rs

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

mixin Fast {
  String mode() {
    return 'fast';
  }
}
mixin Slow {
  String mode() {
    return 'slow';
  }
}
class Runner with Fast, Slow {}
void __vybeMain() {
  __p(Runner().mode());
}

void main() {
  __vybeMain();
  __check('slow');
}
