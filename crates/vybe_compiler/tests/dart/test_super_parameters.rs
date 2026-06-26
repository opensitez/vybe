//! Dart 2.17+ super parameters: Class(super.field), required super,
//! named super parameters, and forwarding to the super constructor.

dart_cases! {
    super_param_forwards_single_int_field => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(super.n);
}
void main() {
  print(Sub(9).n);
}"#,
        ["9"]
    };

    super_param_forwards_string_field => {
        r#"class Base {
  String label;
  Base(this.label);
}
class Sub extends Base {
  Sub(super.label);
}
void main() {
  print(Sub('hi').label);
}"#,
        ["hi"]
    };

    super_param_forwards_bool_field => {
        r#"class Base {
  bool flag;
  Base(this.flag);
}
class Sub extends Base {
  Sub(super.flag);
}
void main() {
  print(Sub(true).flag);
}"#,
        ["true"]
    };

    super_param_forwards_first_of_two_base_fields => {
        r#"class Base {
  int a;
  int b;
  Base(this.a, this.b);
}
class Sub extends Base {
  Sub(super.a, super.b);
}
void main() {
  var s = Sub(2, 3);
  print(s.a + s.b);
}"#,
        ["5"]
    };

    super_param_plus_subclass_field => {
        r#"class Base {
  int x;
  Base(this.x);
}
class Sub extends Base {
  int y;
  Sub(super.x, this.y);
}
void main() {
  var s = Sub(4, 6);
  print(s.x + s.y);
}"#,
        ["10"]
    };

    super_param_subclass_field_before_super => {
        r#"class Base {
  int x;
  Base(this.x);
}
class Sub extends Base {
  int extra;
  Sub(this.extra, super.x);
}
void main() {
  var s = Sub(7, 3);
  print(s.extra + s.x);
}"#,
        ["10"]
    };

    super_param_three_positional_forwards => {
        r#"class Triple {
  int a;
  int b;
  int c;
  Triple(this.a, this.b, this.c);
}
class Child extends Triple {
  Child(super.a, super.b, super.c);
}
void main() {
  print(Child(1, 2, 3).b);
}"#,
        ["2"]
    };

    super_param_named_forwards_base_field => {
        r#"class Base {
  int width;
  Base({this.width = 1});
}
class Sub extends Base {
  Sub({super.width});
}
void main() {
  print(Sub(width: 8).width);
}"#,
        ["8"]
    };

    super_param_named_required_forwards => {
        r#"class Base {
  int id;
  Base({required this.id});
}
class Sub extends Base {
  Sub({required super.id});
}
void main() {
  print(Sub(id: 42).id);
}"#,
        ["42"]
    };

    super_param_named_mixed_with_sub_field => {
        r#"class Base {
  String name;
  Base({required this.name});
}
class Sub extends Base {
  int score;
  Sub({required super.name, this.score = 0});
}
void main() {
  var s = Sub(name: 'Ann', score: 10);
  print('${s.name}:${s.score}');
}"#,
        ["Ann:10"]
    };

    super_param_named_optional_uses_base_default => {
        r#"class Base {
  int level;
  Base({this.level = 5});
}
class Sub extends Base {
  Sub({super.level});
}
void main() {
  print(Sub().level);
}"#,
        ["5"]
    };

    super_param_chain_two_levels => {
        r#"class Root {
  int v;
  Root(this.v);
}
class Mid extends Root {
  Mid(super.v);
}
class Leaf extends Mid {
  Leaf(super.v);
}
void main() {
  print(Leaf(11).v);
}"#,
        ["11"]
    };

    super_param_chain_three_levels_sum => {
        r#"class A {
  int n;
  A(this.n);
}
class B extends A {
  B(super.n);
}
class C extends B {
  C(super.n);
}
void main() {
  print(C(3).n + 1);
}"#,
        ["4"]
    };

    super_param_mid_adds_field_leaf_forwards => {
        r#"class A {
  int base;
  A(this.base);
}
class B extends A {
  int mid;
  B(super.base, this.mid);
}
class C extends B {
  C(super.base, super.mid);
}
void main() {
  print(C(2, 5).mid);
}"#,
        ["5"]
    };

    super_param_with_constructor_body => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  int doubled;
  Sub(super.n) {
    doubled = n * 2;
  }
}
void main() {
  print(Sub(6).doubled);
}"#,
        ["12"]
    };

    super_param_initializer_list_and_body => {
        r#"class Base {
  int x;
  Base(this.x);
}
class Sub extends Base {
  int y;
  Sub(super.x, int add) : y = add {
    y = y + x;
  }
}
void main() {
  print(Sub(3, 4).y);
}"#,
        ["7"]
    };

    super_param_expression_not_allowed_uses_literal => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(super.n);
}
void main() {
  print(Sub(10).n + Sub(5).n);
}"#,
        ["15"]
    };

    super_param_double_field => {
        r#"class Base {
  double rate;
  Base(this.rate);
}
class Sub extends Base {
  Sub(super.rate);
}
void main() {
  print(Sub(2.5).rate + 0.5);
}"#,
        ["3.0"]
    };

    super_param_nullable_int_accepts_null => {
        r#"class Base {
  int? maybe;
  Base(this.maybe);
}
class Sub extends Base {
  Sub(super.maybe);
}
void main() {
  print(Sub(null).maybe == null);
}"#,
        ["true"]
    };

    super_param_nullable_int_accepts_value => {
        r#"class Base {
  int? maybe;
  Base(this.maybe);
}
class Sub extends Base {
  Sub(super.maybe);
}
void main() {
  print(Sub(7).maybe);
}"#,
        ["7"]
    };

    super_param_list_field_length => {
        r#"class Base {
  List<int> items;
  Base(this.items);
}
class Sub extends Base {
  Sub(super.items);
}
void main() {
  print(Sub([1, 2, 3]).items.length);
}"#,
        ["3"]
    };

    super_param_map_field_lookup => {
        r#"class Base {
  Map<String, int> data;
  Base(this.data);
}
class Sub extends Base {
  Sub(super.data);
}
void main() {
  print(Sub({'a': 1}).data['a']);
}"#,
        ["1"]
    };

    super_param_with_base_default_constructor_value => {
        r#"class Base {
  int count = 0;
  Base(this.count);
}
class Sub extends Base {
  Sub(super.count);
}
void main() {
  print(Sub(99).count);
}"#,
        ["99"]
    };

    super_param_subclass_named_constructor => {
        r#"class Base {
  int x;
  int y;
  Base(this.x, this.y);
}
class Sub extends Base {
  Sub(super.x, super.y);
  Sub.origin() : super(0, 0);
}
void main() {
  print(Sub.origin().x + Sub.origin().y);
}"#,
        ["0"]
    };

    super_param_subclass_named_alternate => {
        r#"class Base {
  int w;
  int h;
  Base(this.w, this.h);
}
class Rect extends Base {
  Rect(super.w, super.h);
  Rect.square(int side) : super(side, side);
}
void main() {
  print(Rect.square(4).w + Rect.square(4).h);
}"#,
        ["8"]
    };

    super_param_base_with_getter_read => {
        r#"class Base {
  int n;
  Base(this.n);
  int doubled() {
    return n * 2;
  }
}
class Sub extends Base {
  Sub(super.n);
}
void main() {
  print(Sub(5).doubled());
}"#,
        ["10"]
    };

    super_param_inherited_method_after_forward => {
        r#"class Base {
  String tag;
  Base(this.tag);
  String label() {
    return 'base:$tag';
  }
}
class Sub extends Base {
  Sub(super.tag);
}
void main() {
  print(Sub('x').label());
}"#,
        ["base:x"]
    };

    super_param_override_method_still_works => {
        r#"class Base {
  int n;
  Base(this.n);
  int val() {
    return n;
  }
}
class Sub extends Base {
  Sub(super.n);
  int val() {
    return super.val() + 1;
  }
}
void main() {
  print(Sub(4).val());
}"#,
        ["5"]
    };

    super_param_two_named_both_required => {
        r#"class Base {
  int a;
  int b;
  Base({required this.a, required this.b});
}
class Sub extends Base {
  Sub({required super.a, required super.b});
}
void main() {
  print(Sub(a: 3, b: 4).a + Sub(a: 3, b: 4).b);
}"#,
        ["7"]
    };

    super_param_one_named_one_positional => {
        r#"class Base {
  int x;
  String name;
  Base(this.x, {this.name = 'anon'});
}
class Sub extends Base {
  Sub(super.x, {super.name});
}
void main() {
  print(Sub(1, name: 'bob').name);
}"#,
        ["bob"]
    };

    super_param_optional_positional_super_last => {
        r#"class Base {
  int a;
  int b;
  Base(this.a, [this.b = 0]);
}
class Sub extends Base {
  Sub(super.a, [super.b]);
}
void main() {
  print(Sub(5).b);
}"#,
        ["0"]
    };

    super_param_optional_positional_super_provided => {
        r#"class Base {
  int a;
  int b;
  Base(this.a, [this.b = 0]);
}
class Sub extends Base {
  Sub(super.a, [super.b]);
}
void main() {
  print(Sub(5, 9).b);
}"#,
        ["9"]
    };

    super_param_four_fields_product => {
        r#"class Base {
  int w;
  int h;
  Base(this.w, this.h);
}
class Sized extends Base {
  Sized(super.w, super.h);
  int area() {
    return w * h;
  }
}
void main() {
  print(Sized(3, 4).area());
}"#,
        ["12"]
    };

    super_param_string_concat_in_method => {
        r#"class Base {
  String first;
  String last;
  Base(this.first, this.last);
}
class Person extends Base {
  Person(super.first, super.last);
  String full() {
    return '$first $last';
  }
}
void main() {
  print(Person('Ada', 'Lovelace').full());
}"#,
        ["Ada Lovelace"]
    };

    super_param_negative_int_preserved => {
        r#"class Base {
  int delta;
  Base(this.delta);
}
class Sub extends Base {
  Sub(super.delta);
}
void main() {
  print(Sub(-8).delta + 10);
}"#,
        ["2"]
    };

    super_param_zero_value => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(super.n);
}
void main() {
  print(Sub(0).n == 0);
}"#,
        ["true"]
    };

    super_param_large_int => {
        r#"class Base {
  int big;
  Base(this.big);
}
class Sub extends Base {
  Sub(super.big);
}
void main() {
  print(Sub(1000000).big);
}"#,
        ["1000000"]
    };

    super_param_multiple_instances_independent => {
        r#"class Base {
  int id;
  Base(this.id);
}
class Sub extends Base {
  Sub(super.id);
}
void main() {
  var a = Sub(1);
  var b = Sub(2);
  print(a.id + b.id);
}"#,
        ["3"]
    };

    super_param_field_mutation_after_construct => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(super.n);
}
void main() {
  var s = Sub(1);
  s.n = 20;
  print(s.n);
}"#,
        ["20"]
    };

    super_param_with_base_method_using_field => {
        r#"class Counter {
  int count;
  Counter(this.count);
  void inc() {
    count = count + 1;
  }
}
class StepCounter extends Counter {
  StepCounter(super.count);
}
void main() {
  var c = StepCounter(10);
  c.inc();
  print(c.count);
}"#,
        ["11"]
    };

    super_param_base_two_named_sub_adds_positional => {
        r#"class Config {
  int port;
  String host;
  Config({this.port = 80, this.host = 'localhost'});
}
class AppConfig extends Config {
  String app;
  AppConfig(this.app, {super.port, super.host});
}
void main() {
  var c = AppConfig('vybe', port: 8080);
  print('${c.app}:${c.port}');
}"#,
        ["vybe:8080"]
    };

    super_param_deep_hierarchy_preserves_type => {
        r#"class Entity {
  int id;
  Entity(this.id);
}
class Model extends Entity {
  Model(super.id);
}
class User extends Model {
  User(super.id);
}
void main() {
  print(User(7) is Entity);
}"#,
        ["true"]
    };

    super_param_supertype_assignment => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(super.n);
}
void main() {
  Base b = Sub(13);
  print(b.n);
}"#,
        ["13"]
    };

    super_param_empty_string => {
        r#"class Base {
  String text;
  Base(this.text);
}
class Sub extends Base {
  Sub(super.text);
}
void main() {
  print(Sub('').text.length);
}"#,
        ["0"]
    };

    super_param_bool_false => {
        r#"class Base {
  bool ok;
  Base(this.ok);
}
class Sub extends Base {
  Sub(super.ok);
}
void main() {
  print(Sub(false).ok);
}"#,
        ["false"]
    };
}
