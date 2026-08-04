// vybe-test: dart/async_generators/async_star_with_transform
// origin: languages/dart/tests/dart/test_async_generators.rs

Stream<String> messages() async* {
  var items = ['a', 'b', 'c'];
  for (var item in items) { yield item.toUpperCase(); }
}

void main() {}
