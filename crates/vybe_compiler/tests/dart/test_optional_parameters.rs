//! Optional positional parameters: bracket syntax, default values, trailing
//! optionals, and mixing with required positional parameters.

dart_cases! {
    single_optional_uses_default_when_omitted => {
        r#"void greet([String name = 'World']) {
  print('Hello $name');
}
void main() {
  greet();
}"#,
        ["Hello World"]
    };

    single_optional_overrides_default => {
        r#"void greet([String name = 'World']) {
  print('Hello $name');
}
void main() {
  greet('Dart');
}"#,
        ["Hello Dart"]
    };

    single_optional_int_default => {
        r#"void show([int n = 0]) {
  print(n);
}
void main() {
  show();
}"#,
        ["0"]
    };

    single_optional_int_provided => {
        r#"void show([int n = 0]) {
  print(n);
}
void main() {
  show(99);
}"#,
        ["99"]
    };

    two_optional_both_defaults => {
        r#"void pair([int a = 1, int b = 2]) {
  print('$a,$b');
}
void main() {
  pair();
}"#,
        ["1,2"]
    };

    two_optional_first_provided => {
        r#"void pair([int a = 1, int b = 2]) {
  print('$a,$b');
}
void main() {
  pair(10);
}"#,
        ["10,2"]
    };

    two_optional_both_provided => {
        r#"void pair([int a = 1, int b = 2]) {
  print('$a,$b');
}
void main() {
  pair(10, 20);
}"#,
        ["10,20"]
    };

    required_then_one_optional_default => {
        r#"void log(String level, [String detail = 'none']) {
  print('$level:$detail');
}
void main() {
  log('INFO');
}"#,
        ["INFO:none"]
    };

    required_then_one_optional_set => {
        r#"void log(String level, [String detail = 'none']) {
  print('$level:$detail');
}
void main() {
  log('INFO', 'boot');
}"#,
        ["INFO:boot"]
    };

    required_then_two_optionals_defaults => {
        r#"void fmt(String prefix, [String mid = '-', String suffix = '!']) {
  print('$prefix$mid$suffix');
}
void main() {
  fmt('A');
}"#,
        ["A-!"]
    };

    required_then_two_optionals_first_set => {
        r#"void fmt(String prefix, [String mid = '-', String suffix = '!']) {
  print('$prefix$mid$suffix');
}
void main() {
  fmt('A', '+');
}"#,
        ["A+!"]
    };

    required_then_two_optionals_both_set => {
        r#"void fmt(String prefix, [String mid = '-', String suffix = '!']) {
  print('$prefix$mid$suffix');
}
void main() {
  fmt('A', '+', '?');
}"#,
        ["A+?"]
    };

    optional_bool_default_false => {
        r#"void flag([bool on = false]) {
  print(on);
}
void main() {
  flag();
}"#,
        ["false"]
    };

    optional_bool_set_true => {
        r#"void flag([bool on = false]) {
  print(on);
}
void main() {
  flag(true);
}"#,
        ["true"]
    };

    optional_double_default => {
        r#"void scale([double factor = 1.0]) {
  print(factor * 5);
}
void main() {
  scale();
}"#,
        ["5.0"]
    };

    optional_double_override => {
        r#"void scale([double factor = 1.0]) {
  print(factor * 5);
}
void main() {
  scale(2.0);
}"#,
        ["10.0"]
    };

    three_trailing_optionals_none_given => {
        r#"void seq([int a = 0, int b = 0, int c = 0]) {
  print('$a$b$c');
}
void main() {
  seq();
}"#,
        ["000"]
    };

    three_trailing_optionals_one_given => {
        r#"void seq([int a = 0, int b = 0, int c = 0]) {
  print('$a$b$c');
}
void main() {
  seq(1);
}"#,
        ["100"]
    };

    three_trailing_optionals_two_given => {
        r#"void seq([int a = 0, int b = 0, int c = 0]) {
  print('$a$b$c');
}
void main() {
  seq(1, 2);
}"#,
        ["120"]
    };

    three_trailing_optionals_all_given => {
        r#"void seq([int a = 0, int b = 0, int c = 0]) {
  print('$a$b$c');
}
void main() {
  seq(1, 2, 3);
}"#,
        ["123"]
    };

    optional_string_empty_default => {
        r#"void title([String text = '']) {
  print(text.isEmpty ? 'empty' : text);
}
void main() {
  title();
}"#,
        ["empty"]
    };

    optional_string_nonempty => {
        r#"void title([String text = '']) {
  print(text.isEmpty ? 'empty' : text);
}
void main() {
  title('page');
}"#,
        ["page"]
    };

    optional_list_default_const => {
        r#"void count([List<int> items = const [0]]) {
  print(items.length);
}
void main() {
  count();
}"#,
        ["1"]
    };

    optional_list_override => {
        r#"void count([List<int> items = const [0]]) {
  print(items.length);
}
void main() {
  count([1, 2, 3]);
}"#,
        ["3"]
    };

    optional_param_used_in_arithmetic => {
        r#"int bump([int step = 1]) => 10 + step;
void main() {
  print(bump());
  print(bump(5));
}"#,
        ["11", "15"]
    };

    optional_param_in_class_method => {
        r#"class Greeter {
  void say([String word = 'hi']) {
    print(word);
  }
}
void main() {
  Greeter().say();
  Greeter().say('yo');
}"#,
        ["hi", "yo"]
    };

    optional_param_constructor => {
        r#"class Box {
  final int size;
  Box([this.size = 1]);
}
void main() {
  print(Box().size);
  print(Box(9).size);
}"#,
        ["1", "9"]
    };

    optional_param_with_required_prefix => {
        r#"void repeat(String ch, [int times = 3]) {
  print(ch * times);
}
void main() {
  repeat('x');
  repeat('y', 2);
}"#,
        ["xxx", "yy"]
    };

    optional_nullable_string_default_null => {
        r#"void maybe([String? tag = null]) {
  print(tag ?? 'none');
}
void main() {
  maybe();
}"#,
        ["none"]
    };

    optional_nullable_string_provided => {
        r#"void maybe([String? tag = null]) {
  print(tag ?? 'none');
}
void main() {
  maybe('ok');
}"#,
        ["ok"]
    };

    optional_zero_default_not_confused_with_missing => {
        r#"void show([int n = 0]) {
  print(n == 0 ? 'zero' : '$n');
}
void main() {
  show(0);
}"#,
        ["zero"]
    };

    optional_negative_default => {
        r#"void offset([int delta = -1]) {
  print(10 + delta);
}
void main() {
  offset();
}"#,
        ["9"]
    };

    optional_negative_override => {
        r#"void offset([int delta = -1]) {
  print(10 + delta);
}
void main() {
  offset(3);
}"#,
        ["13"]
    };
}
