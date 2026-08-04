// vybe-test: dart/factory_constructors_deep/factory_from_json_string_to_int_parse
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

class Score {
  int points;
  Score._(this.points);
  factory Score.fromJson(Map<String, dynamic> json) {
    return Score._(int.parse(json['points']));
  }
}
void __vybeMain() {
  __p(Score.fromJson({'points': '42'}).points);
}

void main() {
  __vybeMain();
  __check('42');
}
