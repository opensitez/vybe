// vybe-test: dart/classes_advanced/implements_multiple
// origin: languages/dart/tests/dart/test_classes_advanced.rs

abstract class Readable { String read(); }
abstract class Writable { void write(String s); }
class File implements Readable, Writable {
  String read() => '';
  void write(String s) {}
}

void main() {}
