// vybe-test: dart/generics/generic_method_on_class
// origin: languages/dart/tests/dart/test_generics.rs

class Converter<T> {
  List<T> items;
  Converter(this.items);
  List<R> convert<R>(R Function(T) fn) => items.map(fn).toList();
}

void main() {}
