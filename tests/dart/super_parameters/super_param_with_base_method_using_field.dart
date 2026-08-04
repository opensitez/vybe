// vybe-test: dart/super_parameters/super_param_with_base_method_using_field
// origin: languages/dart/tests/dart/test_super_parameters.rs

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

class Counter {
  int count;
  Counter(this.count);
  void inc() {
    count = count + 1;
  }
}
class StepCounter extends Counter {
  StepCounter(super.count);
}
void __vybeMain() {
  var c = StepCounter(10);
  c.inc();
  __p(c.count);
}

void main() {
  __vybeMain();
  __check('11');
}
