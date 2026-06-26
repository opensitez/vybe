//! Core enum declarations, index/name properties, values list, and enhanced enums.

dart_cases! {
    simple_enum_first_member_index_zero => {
        r#"enum Color { red, green, blue }
void main() {
  print(Color.red.index);
}"#,
        ["0"]
    };

    simple_enum_middle_member_index => {
        r#"enum Color { red, green, blue }
void main() {
  print(Color.green.index);
}"#,
        ["1"]
    };

    simple_enum_last_member_index => {
        r#"enum Color { red, green, blue }
void main() {
  print(Color.blue.index);
}"#,
        ["2"]
    };

    enum_name_property_first_member => {
        r#"enum Direction { north, south, east, west }
void main() {
  print(Direction.north.name);
}"#,
        ["north"]
    };

    enum_name_property_last_member => {
        r#"enum Direction { north, south, east, west }
void main() {
  print(Direction.west.name);
}"#,
        ["west"]
    };

    enum_values_list_length => {
        r#"enum Size { small, medium, large }
void main() {
  print(Size.values.length);
}"#,
        ["3"]
    };

    enum_values_first_element_index => {
        r#"enum Size { small, medium, large }
void main() {
  print(Size.values[0].index);
}"#,
        ["0"]
    };

    enum_values_second_element_name => {
        r#"enum Size { small, medium, large }
void main() {
  print(Size.values[1].name);
}"#,
        ["medium"]
    };

    enum_values_last_element_name => {
        r#"enum Size { small, medium, large }
void main() {
  print(Size.values[2].name);
}"#,
        ["large"]
    };

    enum_equality_same_member => {
        r#"enum Status { active, inactive }
void main() {
  var s = Status.active;
  print(s == Status.active);
}"#,
        ["true"]
    };

    enum_equality_different_members => {
        r#"enum Status { active, inactive }
void main() {
  print(Status.active == Status.inactive);
}"#,
        ["false"]
    };

    enum_assigned_variable_reports_index => {
        r#"enum Fruit { apple, banana, cherry }
void main() {
  var f = Fruit.banana;
  print(f.index);
}"#,
        ["1"]
    };

    enum_assigned_variable_reports_name => {
        r#"enum Fruit { apple, banana, cherry }
void main() {
  var f = Fruit.cherry;
  print(f.name);
}"#,
        ["cherry"]
    };

    enum_single_member_index_is_zero => {
        r#"enum Solo { only }
void main() {
  print(Solo.only.index);
}"#,
        ["0"]
    };

    enum_single_member_name => {
        r#"enum Solo { only }
void main() {
  print(Solo.only.name);
}"#,
        ["only"]
    };

    enum_single_member_values_length => {
        r#"enum Solo { only }
void main() {
  print(Solo.values.length);
}"#,
        ["1"]
    };

    enum_in_if_true_branch => {
        r#"enum Mode { read, write }
void main() {
  var m = Mode.write;
  if (m == Mode.write) {
    print('writable');
  } else {
    print('readonly');
  }
}"#,
        ["writable"]
    };

    enum_in_if_false_branch => {
        r#"enum Mode { read, write }
void main() {
  var m = Mode.read;
  if (m == Mode.write) {
    print('writable');
  } else {
    print('readonly');
  }
}"#,
        ["readonly"]
    };

    enum_switch_matches_first_case => {
        r#"enum Tier { bronze, silver, gold }
void main() {
  var t = Tier.bronze;
  switch (t) {
    case Tier.bronze:
      print('third');
      break;
    case Tier.silver:
      print('second');
      break;
    default:
      print('first');
  }
}"#,
        ["third"]
    };

    enum_switch_matches_middle_case => {
        r#"enum Tier { bronze, silver, gold }
void main() {
  var t = Tier.silver;
  switch (t) {
    case Tier.bronze:
      print('third');
      break;
    case Tier.silver:
      print('second');
      break;
    default:
      print('first');
  }
}"#,
        ["second"]
    };

    enum_switch_matches_default_case => {
        r#"enum Tier { bronze, silver, gold }
void main() {
  var t = Tier.gold;
  switch (t) {
    case Tier.bronze:
      print('third');
      break;
    case Tier.silver:
      print('second');
      break;
    default:
      print('first');
  }
}"#,
        ["first"]
    };

    enum_values_iterate_sum_indices => {
        r#"enum Digit { zero, one, two }
void main() {
  var sum = 0;
  for (var d in Digit.values) {
    sum += d.index;
  }
  print(sum);
}"#,
        ["3"]
    };

    enum_values_map_to_names_joined => {
        r#"enum Axis { x, y, z }
void main() {
  var names = Axis.values.map((d) => d.name).join('-');
  print(names);
}"#,
        ["x-y-z"]
    };

    enhanced_enum_field_value_first => {
        r#"enum HttpCode {
  ok(200),
  notFound(404);
  final int code;
  const HttpCode(this.code);
}
void main() {
  print(HttpCode.ok.code);
}"#,
        ["200"]
    };

    enhanced_enum_field_value_second => {
        r#"enum HttpCode {
  ok(200),
  notFound(404);
  final int code;
  const HttpCode(this.code);
}
void main() {
  print(HttpCode.notFound.code);
}"#,
        ["404"]
    };

    enhanced_enum_compare_field_in_condition => {
        r#"enum HttpCode {
  ok(200),
  notFound(404);
  final int code;
  const HttpCode(this.code);
}
void main() {
  var c = HttpCode.ok;
  print(c.code == 200);
}"#,
        ["true"]
    };

    enhanced_enum_getter_on_weekend_member => {
        r#"enum Day { monday, tuesday, wednesday, thursday, friday, saturday, sunday;
  bool get isWeekend => this == Day.saturday || this == Day.sunday;
}
void main() {
  print(Day.saturday.isWeekend);
}"#,
        ["true"]
    };

    enhanced_enum_getter_false_on_weekday => {
        r#"enum Day { monday, tuesday, wednesday, thursday, friday, saturday, sunday;
  bool get isWeekend => this == Day.saturday || this == Day.sunday;
}
void main() {
  print(Day.monday.isWeekend);
}"#,
        ["false"]
    };

    enum_four_members_values_length => {
        r#"enum Season { spring, summer, autumn, winter }
void main() {
  print(Season.values.length);
}"#,
        ["4"]
    };

    enum_values_third_element_index => {
        r#"enum Season { spring, summer, autumn, winter }
void main() {
  print(Season.values[2].index);
}"#,
        ["2"]
    };

    enum_values_fourth_element_name => {
        r#"enum Season { spring, summer, autumn, winter }
void main() {
  print(Season.values[3].name);
}"#,
        ["winter"]
    };

    enum_not_equal_operator => {
        r#"enum Op { add, sub }
void main() {
  print(Op.add != Op.sub);
}"#,
        ["true"]
    };

    enum_member_from_values_by_index => {
        r#"enum Level { low, mid, high }
void main() {
  var picked = Level.values[1];
  print(picked.name);
}"#,
        ["mid"]
    };

    enum_printed_name_via_variable => {
        r#"enum Planet { mercury, venus, earth }
void main() {
  var home = Planet.earth;
  print('planet:${home.name}');
}"#,
        ["planet:earth"]
    };

    enum_index_in_arithmetic => {
        r#"enum Step { a, b, c, d }
void main() {
  print(Step.c.index + Step.a.index);
}"#,
        ["2"]
    };

    enum_two_member_values_both_names => {
        r#"enum Binary { off, on }
void main() {
  print(Binary.values[0].name);
  print(Binary.values[1].name);
}"#,
        ["off", "on"]
    };

    enum_enhanced_three_values_sum_codes => {
        r#"enum Priority {
  low(1),
  medium(5),
  high(10);
  final int weight;
  const Priority(this.weight);
}
void main() {
  print(Priority.low.weight + Priority.high.weight);
}"#,
        ["11"]
    };

    enum_enhanced_name_and_code => {
        r#"enum Priority {
  low(1),
  medium(5),
  high(10);
  final int weight;
  const Priority(this.weight);
}
void main() {
  var p = Priority.medium;
  print('${p.name}:${p.weight}');
}"#,
        ["medium:5"]
    };

    enum_identity_same_reference => {
        r#"enum Token { alpha, beta }
void main() {
  var a = Token.alpha;
  var b = Token.alpha;
  print(a == b);
}"#,
        ["true"]
    };

    enum_switch_with_enum_return => {
        r#"enum Shape { circle, square }
String label(Shape s) {
  switch (s) {
    case Shape.circle:
      return 'round';
    default:
      return 'flat';
  }
}
void main() {
  print(label(Shape.square));
}"#,
        ["flat"]
    };
}
