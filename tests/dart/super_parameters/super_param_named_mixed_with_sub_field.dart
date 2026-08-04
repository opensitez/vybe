// vybe-test: dart/super_parameters/super_param_named_mixed_with_sub_field
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

class Base {
  String name;
  Base({required this.name});
}
class Sub extends Base {
  int score;
  Sub({required super.name, this.score = 0});
}
void __vybeMain() {
  var s = Sub(name: 'Ann', score: 10);
  __p('${s.name}:${s.score}');
}

void main() {
  __vybeMain();
  __check('Ann:10');
}
