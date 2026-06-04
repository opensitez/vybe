use super::helpers::{compile_ok, run_prints};

// ── Enum fundamentals ────────────────────────────────────────

#[test]
fn enum_values_list() {
    compile_ok("enum Dir { north, south, east, west } var all = Dir.values;");
}
#[test]
fn enum_index() {
    compile_ok("enum Dir { north, south } var i = Dir.north.index;");
}
#[test]
fn enum_name() {
    compile_ok("enum Dir { north, south } var n = Dir.north.name;");
}

#[test]
fn enum_index_result() {
    let out = run_prints("enum Color { red, green, blue } void main() { print(Color.red.index); }");
    assert_eq!(out, ["0"]);
}

#[test]
fn enum_name_result() {
    let out =
        run_prints("enum Color { red, green, blue } void main() { print(Color.green.name); }");
    assert_eq!(out, ["green"]);
}

#[test]
fn enum_in_switch() {
    let out = run_prints(
        r#"
enum Status { active, inactive, pending }
void main() {
  var s = Status.active;
  switch (s) {
    case Status.active: print('on'); break;
    case Status.inactive: print('off'); break;
    default: print('pending');
  }
}
"#,
    );
    assert_eq!(out, ["on"]);
}

#[test]
fn enum_comparison() {
    let out = run_prints(
        r#"
enum Size { small, medium, large }
void main() {
  var s = Size.medium;
  print(s == Size.medium);
}
"#,
    );
    assert_eq!(out, ["true"]);
}

#[test]
fn enum_in_if() {
    let out = run_prints(
        r#"
enum Direction { up, down, left, right }
void main() {
  var d = Direction.up;
  if (d == Direction.up) { print('going up'); }
}
"#,
    );
    assert_eq!(out, ["going up"]);
}

#[test]
fn enum_values_length() {
    let out = run_prints(
        "enum Season { spring, summer, autumn, winter } void main() { print(Season.values.length); }",
    );
    assert_eq!(out, ["4"]);
}

#[test]
fn enum_assigned_to_var() {
    let out = run_prints(
        "enum Fruit { apple, banana } void main() { var f = Fruit.banana; print(f.index); }",
    );
    assert_eq!(out, ["1"]);
}

#[test]
fn enum_in_list() {
    compile_ok(
        "enum Color { red, green, blue } var palette = [Color.red, Color.green, Color.blue];",
    );
}

#[test]
fn enum_in_map() {
    compile_ok("enum Level { easy, hard } var scores = {Level.easy: 10, Level.hard: 50};");
}

// ── Enhanced enums (Dart 3) ──────────────────────────────────

#[test]
fn enhanced_enum_with_field() {
    compile_ok(
        r#"
enum Planet {
  mercury(3.303e+23, 2.4397e6),
  venus(4.869e+24, 6.0518e6);

  final double mass;
  final double radius;
  const Planet(this.mass, this.radius);
}
"#,
    );
}

#[test]
fn enhanced_enum_with_method() {
    compile_ok(
        r#"
enum Day {
  monday, tuesday, wednesday, thursday, friday, saturday, sunday;

  bool get isWeekend => this == Day.saturday || this == Day.sunday;
}
"#,
    );
}

#[test]
fn enhanced_enum_toString() {
    compile_ok(
        r#"
enum Color {
  red, green, blue;
  String describe() => 'Color.$name';
}
"#,
    );
}

#[test]
fn enhanced_enum_implements() {
    compile_ok(
        r#"
abstract class Describable { String describe(); }
enum Status implements Describable {
  active, inactive;
  String describe() => 'Status is $name';
}
"#,
    );
}

// ── Enum in class fields ──────────────────────────────────────

#[test]
fn enum_as_class_field() {
    compile_ok(
        r#"
enum Role { admin, user, guest }
class User {
  String name;
  Role role;
  User(this.name, this.role);
}
"#,
    );
}

#[test]
fn enum_default_param() {
    compile_ok(
        r#"
enum Level { low, medium, high }
class Alert {
  String msg;
  Level level;
  Alert(this.msg, {this.level = Level.low});
}
"#,
    );
}

// ── Enum iteration ───────────────────────────────────────────

#[test]
fn enum_iterate_values() {
    compile_ok(
        r#"
enum Color { red, green, blue }
void main() {
  for (var c in Color.values) {
    print(c.name);
  }
}
"#,
    );
}

#[test]
fn enum_map_over() {
    compile_ok(
        r#"
enum Priority { low, medium, high }
void main() {
  var names = Priority.values.map((e) => e.name).toList();
}
"#,
    );
}

// ── Enum return from function ────────────────────────────────

#[test]
fn enum_return() {
    compile_ok(
        r#"
enum Result { ok, err }
Result check(int x) { return x > 0 ? Result.ok : Result.err; }
"#,
    );
}

#[test]
fn enum_return_result() {
    let out = run_prints(
        r#"
enum Result { ok, err }
Result check(int x) { return x > 0 ? Result.ok : Result.err; }
void main() { print(check(1).name); }
"#,
    );
    assert_eq!(out, ["ok"]);
}
