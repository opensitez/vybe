// vybe-test: dart/classes_advanced/mixin_with_field
// origin: languages/dart/tests/dart/test_classes_advanced.rs

mixin Timestamped {
  late DateTime createdAt;
  void initTimestamp() { createdAt = DateTime.now(); }
}
class Post with Timestamped { String content; Post(this.content); }

void main() {}
