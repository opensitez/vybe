// vybe-test: dart/covariant_keyword/covariant_param_compare_type_tag
// origin: languages/dart/tests/dart/test_covariant_keyword.rs

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

class Item {}
class Book extends Item {
  String title;
  Book(this.title);
}
class Shelf {
  String label(Item i) {
    return 'item';
  }
}
class BookShelf extends Shelf {
  @override
  String label(covariant Book b) {
    return b.title;
  }
}
void __vybeMain() {
  __p(BookShelf().label(Book('Dart')));
}

void main() {
  __vybeMain();
  __check('Dart');
}
