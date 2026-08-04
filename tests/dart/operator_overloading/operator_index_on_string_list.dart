// vybe-test: dart/operator_overloading/operator_index_on_string_list
// origin: languages/dart/tests/dart/test_operator_overloading.rs

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

class Words {
  List<String> items;
  Words(this.items);
  String operator [](int i) {
    return items[i];
  }
}
void __vybeMain() {
  __p(Words(['a', 'b'])[1]);
}

void main() {
  __vybeMain();
  __check('b');
}
