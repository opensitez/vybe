// vybe-test: dart/classes_advanced/extension_on_list
// origin: languages/dart/tests/dart/test_classes_advanced.rs

extension ListExt<T> on List<T> { T? get secondOrNull => length > 1 ? this[1] : null; }

void main() {}
