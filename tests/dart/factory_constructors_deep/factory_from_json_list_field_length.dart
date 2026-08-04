// vybe-test: dart/factory_constructors_deep/factory_from_json_list_field_length
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

class Batch {
  List<int> ids;
  Batch._(this.ids);
  factory Batch.fromJson(Map<String, dynamic> json) {
    var raw = json['ids'] as List;
    return Batch._(raw.cast<int>());
  }
}
void __vybeMain() {
  var b = Batch.fromJson({'ids': [1, 2, 3]});
  __p(b.ids.length);
}

void main() {
  __vybeMain();
  __check('3');
}
