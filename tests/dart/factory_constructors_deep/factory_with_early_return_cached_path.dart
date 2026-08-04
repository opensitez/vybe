// vybe-test: dart/factory_constructors_deep/factory_with_early_return_cached_path
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

class Blob {
  static Blob? _empty;
  int size;
  Blob._(this.size);
  factory Blob.empty() {
    if (_empty != null) {
      return _empty!;
    }
    _empty = Blob._(0);
    return _empty!;
  }
}
void __vybeMain() {
  __p(Blob.empty().size);
}

void main() {
  __vybeMain();
  __check('0');
}
