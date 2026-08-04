// vybe-test: dart/abstract_members/abstract_class_with_final_field_in_sub
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

abstract class IdEntity {
  String get id;
}
class Record extends IdEntity {
  final String _id;
  Record(this._id);
  String get id => _id;
}
void __vybeMain() {
  __p(Record('r42').id);
}

void main() {
  __vybeMain();
  __check('r42');
}
