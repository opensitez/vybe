// vybe-test: dart/classes_advanced/implements_with_extends
// origin: languages/dart/tests/dart/test_classes_advanced.rs

abstract class Serializable { String serialize(); }
class Base { int id; Base(this.id); }
class Entity extends Base implements Serializable {
  Entity(int id) : super(id);
  String serialize() => '{"id": $id}';
}

void main() {}
