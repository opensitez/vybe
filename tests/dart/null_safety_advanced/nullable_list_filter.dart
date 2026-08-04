// vybe-test: dart/null_safety_advanced/nullable_list_filter
// origin: languages/dart/tests/dart/test_null_safety_advanced.rs

List<String?> list = ['a', null, 'b']; var nonNull = list.where((e) => e != null).toList();

void main() {}
