//! super.method calls and super(...) constructor initializers.

dart_cases! {
    super_constructor_passes_single_int => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(int v) : super(v);
}
void main() {
  print(Sub(7).n);
}"#,
        ["7"]
    };

    super_constructor_passes_two_args => {
        r#"class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
}
class Tagged extends Pair {
  Tagged(int x, int y) : super(x, y);
}
void main() {
  var t = Tagged(2, 5);
  print(t.a + t.b);
}"#,
        ["7"]
    };

    super_initializer_with_expression => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  Sub(int x) : super(x * 2);
}
void main() {
  print(Sub(4).n);
}"#,
        ["8"]
    };

    super_initializer_before_constructor_body => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Sub extends Base {
  int extra;
  Sub(int a, int b) : super(a) {
    extra = b;
  }
}
void main() {
  var s = Sub(3, 10);
  print(s.n + s.extra);
}"#,
        ["13"]
    };

    super_method_in_overridden_string_method => {
        r#"class Greeter {
  String hello() {
    return 'hi';
  }
}
class Loud extends Greeter {
  String hello() {
    return super.hello() + '!';
  }
}
void main() {
  print(Loud().hello());
}"#,
        ["hi!"]
    };

    super_method_appends_suffix => {
        r#"class Base {
  String tag() {
    return 'core';
  }
}
class Wrap extends Base {
  String tag() {
    return super.tag() + '-wrap';
  }
}
void main() {
  print(Wrap().tag());
}"#,
        ["core-wrap"]
    };

    super_method_prefixes_result => {
        r#"class Base {
  String code() {
    return 'x';
  }
}
class Prefixed extends Base {
  String code() {
    return 'p-' + super.code();
  }
}
void main() {
  print(Prefixed().code());
}"#,
        ["p-x"]
    };

    super_method_adds_numeric_offset => {
        r#"class Counter {
  int value() {
    return 10;
  }
}
class Boost extends Counter {
  int value() {
    return super.value() + 5;
  }
}
void main() {
  print(Boost().value());
}"#,
        ["15"]
    };

    super_method_called_from_helper => {
        r#"class Base {
  int baseVal() {
    return 4;
  }
}
class Child extends Base {
  int combined() {
    return super.baseVal() + 1;
  }
}
void main() {
  print(Child().combined());
}"#,
        ["5"]
    };

    super_constructor_with_string_arg => {
        r#"class Named {
  String name;
  Named(this.name);
}
class Alias extends Named {
  Alias(String n) : super(n);
}
void main() {
  print(Alias('dart').name);
}"#,
        ["dart"]
    };

    super_in_three_level_chain => {
        r#"class A {
  int depth() {
    return 1;
  }
}
class B extends A {
  int depth() {
    return super.depth() + 1;
  }
}
class C extends B {
  int depth() {
    return super.depth() + 1;
  }
}
void main() {
  print(C().depth());
}"#,
        ["3"]
    };

    super_constructor_named_to_base => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
}
class Origin extends Point {
  Origin() : super(0, 0);
}
void main() {
  var o = Origin();
  print(o.x + o.y);
}"#,
        ["0"]
    };

    super_method_doubles_base_int => {
        r#"class Engine {
  int power() {
    return 50;
  }
}
class Turbo extends Engine {
  int power() {
    return super.power() * 2;
  }
}
void main() {
  print(Turbo().power());
}"#,
        ["100"]
    };

    super_initializer_uses_field_expression => {
        r#"class Holder {
  int size;
  Holder(this.size);
}
class Scaled extends Holder {
  Scaled(int n) : super(n + n);
}
void main() {
  print(Scaled(3).size);
}"#,
        ["6"]
    };

    super_method_in_getter_override => {
        r#"class Base {
  int get val {
    return 1;
  }
}
class Child extends Base {
  int get val {
    return super.val + 9;
  }
}
void main() {
  print(Child().val);
}"#,
        ["10"]
    };

    super_constructor_passes_zero => {
        r#"class Slot {
  int id;
  Slot(this.id);
}
class Empty extends Slot {
  Empty() : super(0);
}
void main() {
  print(Empty().id);
}"#,
        ["0"]
    };

    super_method_concat_three_parts => {
        r#"class Part {
  String mid() {
    return 'b';
  }
}
class Full extends Part {
  String mid() {
    return 'a' + super.mid() + 'c';
  }
}
void main() {
  print(Full().mid());
}"#,
        ["abc"]
    };

    super_called_from_subclass_constructor_body => {
        r#"class Logger {
  String msg;
  Logger(this.msg);
}
class Audit extends Logger {
  Audit(String base) : super(base) {
    msg = msg + '-audit';
  }
}
void main() {
  print(Audit('log').msg);
}"#,
        ["log-audit"]
    };

    super_method_with_conditional_extension => {
        r#"class Base {
  int score(bool bonus) {
    return bonus ? 2 : 1;
  }
}
class Plus extends Base {
  int score(bool bonus) {
    return super.score(bonus) + 10;
  }
}
void main() {
  print(Plus().score(true));
}"#,
        ["12"]
    };

    super_initializer_with_negative_literal => {
        r#"class Axis {
  int v;
  Axis(this.v);
}
class Neg extends Axis {
  Neg() : super(-1);
}
void main() {
  print(Neg().v);
}"#,
        ["-1"]
    };

    super_method_returns_bool_from_base => {
        r#"class Check {
  bool ok() {
    return true;
  }
}
class Verify extends Check {
  bool ok() {
    return super.ok();
  }
}
void main() {
  print(Verify().ok());
}"#,
        ["true"]
    };

    super_constructor_forwards_named_param_value => {
        r#"class User {
  String role;
  User(this.role);
}
class Admin extends User {
  Admin() : super('admin');
}
void main() {
  print(Admin().role);
}"#,
        ["admin"]
    };

    super_method_invoked_twice_in_override => {
        r#"class Dup {
  int unit() {
    return 1;
  }
}
class Double extends Dup {
  int unit() {
    return super.unit() + super.unit();
  }
}
void main() {
  print(Double().unit());
}"#,
        ["2"]
    };

    super_in_mixin_on_class_hierarchy => {
        r#"class Root {
  int rootVal() {
    return 2;
  }
}
class Mid extends Root {
  int rootVal() {
    return super.rootVal() + 3;
  }
}
void main() {
  print(Mid().rootVal());
}"#,
        ["5"]
    };

    super_initializer_with_sum_of_literals => {
        r#"class SumBase {
  int total;
  SumBase(this.total);
}
class SumChild extends SumBase {
  SumChild() : super(10 + 5);
}
void main() {
  print(SumChild().total);
}"#,
        ["15"]
    };

    super_method_after_local_var_in_override => {
        r#"class Base {
  String build() {
    return 'base';
  }
}
class Decorated extends Base {
  String build() {
    var prefix = '>>';
    return prefix + super.build();
  }
}
void main() {
  print(Decorated().build());
}"#,
        [">>>base"]
    };

    super_constructor_with_this_shorthand_sibling => {
        r#"class Node {
  int value;
  Node(this.value);
}
class Leaf extends Node {
  int tag;
  Leaf(this.tag, int v) : super(v);
}
void main() {
  var leaf = Leaf(1, 8);
  print(leaf.value);
}"#,
        ["8"]
    };

    super_method_subtracts_from_base => {
        r#"class Wallet {
  int balance() {
    return 20;
  }
}
class Fee extends Wallet {
  int balance() {
    return super.balance() - 5;
  }
}
void main() {
  print(Fee().balance());
}"#,
        ["15"]
    };

    super_initializer_sets_base_before_sub_field => {
        r#"class Base {
  int n;
  Base(this.n);
}
class PairSub extends Base {
  int m;
  PairSub(int a, int b) : super(a), m = b;
}
void main() {
  var p = PairSub(2, 3);
  print(p.n + p.m);
}"#,
        ["5"]
    };

    super_method_used_in_arithmetic_expression => {
        r#"class Meter {
  int read() {
    return 6;
  }
}
class Adjusted extends Meter {
  int read() {
    return super.read() * 2 + 1;
  }
}
void main() {
  print(Adjusted().read());
}"#,
        ["13"]
    };
}
