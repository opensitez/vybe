// vybe-test: dart/enums_advanced/enum_default_param
// origin: languages/dart/tests/dart/test_enums_advanced.rs

enum Level { low, medium, high }
class Alert {
  String msg;
  Level level;
  Alert(this.msg, {this.level = Level.low});
}

void main() {}
