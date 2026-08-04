// vybe-test: dart/no_such_method/no_such_method_sums_positional_arguments
// origin: languages/dart/tests/dart/test_no_such_method.rs

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

class Adder {
  @override
  dynamic noSuchMethod(Invocation inv) {
    var total = 0;
    for (var a in inv.positionalArguments) {
      total = total + (a as int);
    }
    return total;
  }
}
void __vybeMain() {
  dynamic a = Adder();
  __p(a.add(10, 20, 30));
}

void main() {
  __vybeMain();
  __check('60');
}
