// vybe-test: dart/custom_iterator/custom_iterable_enumerate_style_manual
// origin: languages/dart/tests/dart/test_custom_iterator.rs

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

class Tagged extends IterableBase<String> {
  final List<String> items;
  Tagged(this.items);
  @override
  Iterator<String> get iterator => TaggedIterator(items);
}
class TaggedIterator implements Iterator<String> {
  int _i = -1;
  final List<String> items;
  TaggedIterator(this.items);
  @override
  String get current => '${_i}:${items[_i]}';
  @override
  bool moveNext() {
    if (_i + 1 < items.length) {
      _i = _i + 1;
      return true;
    }
    return false;
  }
}
void __vybeMain() {
  __p(Tagged(['a', 'b']).join('|'));
}

void main() {
  __vybeMain();
  __check('0:a|1:b');
}
