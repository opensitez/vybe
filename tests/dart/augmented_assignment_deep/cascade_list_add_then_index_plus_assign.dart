// vybe-test: dart/augmented_assignment_deep/cascade_list_add_then_index_plus_assign
// origin: languages/dart/tests/dart/test_augmented_assignment_deep.rs

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

class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void __vybeMain() {
  var box = Box();
  box..items.add(1)..items[0] += 4;
  __p(box.items[0]);
}

void main() {
  __vybeMain();
  __check('5');
}
