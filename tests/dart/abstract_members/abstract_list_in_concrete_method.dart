// vybe-test: dart/abstract_members/abstract_list_in_concrete_method
// origin: languages/dart/tests/dart/test_abstract_members.rs

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

abstract class Collector {
  List<String> items = [];
  void collect(String s) {
    items.add(s);
  }
  int total();
}
class SimpleCollector extends Collector {
  int total() {
    return items.length;
  }
}
void __vybeMain() {
  var c = SimpleCollector();
  c.collect('a');
  c.collect('b');
  __p(c.total());
}

void main() {
  __vybeMain();
  __check('2');
}
