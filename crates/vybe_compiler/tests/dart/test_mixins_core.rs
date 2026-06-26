//! Dart mixin composition: with, on constraints, and multiple mixins.

dart_cases! {
    basic_mixin_method_available_on_class => {
        r#"mixin Greet {
  String hello() {
    return 'hi';
  }
}
class Person with Greet {}
void main() {
  print(Person().hello());
}"#,
        ["hi"]
    };

    mixin_field_accessible_on_host_class => {
        r#"mixin Tagged {
  String tag = 'm';
}
class Item with Tagged {}
void main() {
  print(Item().tag);
}"#,
        ["m"]
    };

    mixin_method_can_mutate_mixin_field => {
        r#"mixin Counter {
  int n = 0;
  void bump() {
    n = n + 1;
  }
}
class Box with Counter {}
void main() {
  var b = Box();
  b.bump();
  print(b.n);
}"#,
        ["1"]
    };

    with_single_mixin_on_class_with_own_field => {
        r#"mixin Named {
  String label() {
    return 'named';
  }
}
class Widget {
  int id = 1;
  Widget();
}
class NamedWidget extends Widget with Named {}
void main() {
  var w = NamedWidget();
  print(w.label());
}"#,
        ["named"]
    };

    multiple_mixins_both_methods_available => {
        r#"mixin A {
  int a() {
    return 1;
  }
}
mixin B {
  int b() {
    return 2;
  }
}
class Both with A, B {}
void main() {
  var x = Both();
  print(x.a() + x.b());
}"#,
        ["3"]
    };

    multiple_mixins_second_method_invoked => {
        r#"mixin Fly {
  String mode() {
    return 'air';
  }
}
mixin Swim {
  String mode() {
    return 'water';
  }
}
class Duck with Fly, Swim {}
void main() {
  print(Duck().mode());
}"#,
        ["water"]
    };

    mixin_on_constraint_requires_supertype => {
        r#"class Animal {
  String kind = 'animal';
}
mixin Pet on Animal {
  String care() {
    return 'feed';
  }
}
class Dog extends Animal with Pet {}
void main() {
  print(Dog().care());
}"#,
        ["feed"]
    };

    mixin_on_reads_superclass_field => {
        r#"class Vehicle {
  int wheels = 4;
}
mixin Wheeled on Vehicle {
  int countWheels() {
    return wheels;
  }
}
class Car extends Vehicle with Wheeled {}
void main() {
  print(Car().countWheels());
}"#,
        ["4"]
    };

    mixin_on_with_subclass_constructor => {
        r#"class Base {
  int n;
  Base(this.n);
}
mixin Extra on Base {
  int doubled() {
    return n * 2;
  }
}
class Sub extends Base with Extra {
  Sub(int v) : super(v);
}
void main() {
  print(Sub(3).doubled());
}"#,
        ["6"]
    };

    mixin_overrides_method_from_superclass => {
        r#"class Base {
  String talk() {
    return 'base';
  }
}
mixin Loud on Base {
  String talk() {
    return 'loud';
  }
}
class Speaker extends Base with Loud {}
void main() {
  print(Speaker().talk());
}"#,
        ["loud"]
    };

    later_mixin_wins_over_earlier_for_same_name => {
        r#"mixin First {
  int pick() {
    return 1;
  }
}
mixin Second {
  int pick() {
    return 2;
  }
}
class Combo with First, Second {}
void main() {
  print(Combo().pick());
}"#,
        ["2"]
    };

    mixin_with_extends_superclass_method_still_visible => {
        r#"class Root {
  int rootVal() {
    return 10;
  }
}
mixin Branch {
  int branchVal() {
    return 1;
  }
}
class Tree extends Root with Branch {}
void main() {
  var t = Tree();
  print(t.rootVal() + t.branchVal());
}"#,
        ["11"]
    };

    mixin_getter_available_on_host => {
        r#"mixin Sized {
  int get size {
    return 5;
  }
}
class Pack with Sized {}
void main() {
  print(Pack().size);
}"#,
        ["5"]
    };

    mixin_setter_mutates_state => {
        r#"mixin Mutable {
  int _v = 0;
  int get v {
    return _v;
  }
  set v(int n) {
    _v = n;
  }
}
class Holder with Mutable {}
void main() {
  var h = Holder();
  h.v = 9;
  print(h.v);
}"#,
        ["9"]
    };

    mixin_arrow_method_body => {
        r#"mixin Double {
  int twice(int n) => n * 2;
}
class Calc with Double {}
void main() {
  print(Calc().twice(6));
}"#,
        ["12"]
    };

    three_mixins_all_methods_callable => {
        r#"mixin M1 {
  int one() {
    return 1;
  }
}
mixin M2 {
  int two() {
    return 2;
  }
}
mixin M3 {
  int three() {
    return 3;
  }
}
class All with M1, M2, M3 {}
void main() {
  var a = All();
  print(a.one() + a.two() + a.three());
}"#,
        ["6"]
    };

    mixin_on_with_multiple_mixins => {
        r#"class Engine {
  int power = 100;
}
mixin Turbo on Engine {
  int boost() {
    return power + 50;
  }
}
mixin Eco on Engine {
  int save() {
    return power - 20;
  }
}
class Motor extends Engine with Turbo, Eco {}
void main() {
  var m = Motor();
  print(m.boost());
}"#,
        ["150"]
    };

    mixin_method_uses_this => {
        r#"mixin Self {
  int id() {
    return 1;
  }
  int same() {
    return this.id();
  }
}
class Host with Self {}
void main() {
  print(Host().same());
}"#,
        ["1"]
    };

    mixin_with_class_own_method_no_conflict => {
        r#"mixin Mix {
  String fromMix() {
    return 'mix';
  }
}
class Host with Mix {
  String fromHost() {
    return 'host';
  }
}
void main() {
  var h = Host();
  print(h.fromHost() + h.fromMix());
}"#,
        ["hostmix"]
    };

    mixin_applied_to_subclass_not_base => {
        r#"class Base {
  String baseId() {
    return 'b';
  }
}
class Mid extends Base {}
mixin Tag {
  String tagId() {
    return 't';
  }
}
class Leaf extends Mid with Tag {}
void main() {
  var l = Leaf();
  print(l.baseId() + l.tagId());
}"#,
        ["bt"]
    };

    mixin_on_abstract_supertype => {
        r#"abstract class Shape {
  int sides = 0;
}
mixin Polygon on Shape {
  int count() {
    return sides;
  }
}
class Tri extends Shape with Polygon {
  Tri() {
    sides = 3;
  }
}
void main() {
  print(Tri().count());
}"#,
        ["3"]
    };

    two_instances_with_mixin_have_separate_state => {
        r#"mixin State {
  int n = 0;
}
class Node with State {}
void main() {
  var a = Node();
  var b = Node();
  a.n = 5;
  print(b.n);
}"#,
        ["0"]
    };

    mixin_calling_method_from_same_mixin => {
        r#"mixin Chain {
  int step1() {
    return 2;
  }
  int step2() {
    return step1() + 3;
  }
}
class Run with Chain {}
void main() {
  print(Run().step2());
}"#,
        ["5"]
    };

    mixin_with_static_host_class_field => {
        r#"class Host {
  static int global = 7;
}
mixin ReadGlobal on Host {
  int read() {
    return global;
  }
}
class App extends Host with ReadGlobal {}
void main() {
  print(App().read());
}"#,
        ["7"]
    };

    mixin_override_with_super_from_on_type => {
        r#"class Base {
  String val() {
    return 'b';
  }
}
mixin Mid on Base {
  String val() {
    return 'm';
  }
}
class End extends Base with Mid {}
void main() {
  print(End().val());
}"#,
        ["m"]
    };

    mixin_combined_with_class_constructor_field => {
        r#"mixin Id {
  int show(int id) {
    return id;
  }
}
class Record with Id {
  int id;
  Record(this.id);
}
void main() {
  var r = Record(4);
  print(r.show(r.id));
}"#,
        ["4"]
    };

    mixin_on_constraint_with_super_call_in_method => {
        r#"class Base {
  int n = 1;
  int baseInc() {
    return n + 1;
  }
}
mixin Wrap on Base {
  int wrapped() {
    return baseInc() + 1;
  }
}
class Sub extends Base with Wrap {}
void main() {
  print(Sub().wrapped());
}"#,
        ["3"]
    };

    four_mixins_linearized_last_wins => {
        r#"mixin A {
  String tag() {
    return 'a';
  }
}
mixin B {
  String tag() {
    return 'b';
  }
}
mixin C {
  String tag() {
    return 'c';
  }
}
mixin D {
  String tag() {
    return 'd';
  }
}
class X with A, B, C, D {}
void main() {
  print(X().tag());
}"#,
        ["d"]
    };

    mixin_void_method_side_effect => {
        r#"mixin Logger {
  int hits = 0;
  void hit() {
    hits = hits + 1;
  }
}
class Target with Logger {}
void main() {
  var t = Target();
  t.hit();
  t.hit();
  print(t.hits);
}"#,
        ["2"]
    };

    mixin_on_with_named_constructor_on_subclass => {
        r#"class Base {
  int v;
  Base(this.v);
}
mixin Scale on Base {
  int scaled() {
    return v * 10;
  }
}
class Sub extends Base with Scale {
  Sub.zero() : super(0);
}
void main() {
  print(Sub.zero().scaled());
}"#,
        ["0"]
    };

    mixin_method_returns_host_type_field_sum => {
        r#"mixin Sum {
  int total(int a, int b) {
    return a + b;
  }
}
class Pair with Sum {
  int x = 2;
  int y = 3;
}
void main() {
  var p = Pair();
  print(p.total(p.x, p.y));
}"#,
        ["5"]
    };

    mixin_applied_without_extends => {
        r#"mixin M {
  int get flag {
    return 1;
  }
}
class Plain with M {}
void main() {
  print(Plain().flag);
}"#,
        ["1"]
    };

    mixin_on_hierarchy_three_levels => {
        r#"class A {
  int a() {
    return 1;
  }
}
class B extends A {}
mixin C on B {
  int c() {
    return 10;
  }
}
class D extends B with C {}
void main() {
  var d = D();
  print(d.a() + d.c());
}"#,
        ["11"]
    };

    mixin_with_parameterized_method => {
        r#"mixin Format {
  String pad(String s, int n) {
    return s + n.toString();
  }
}
class Tool with Format {}
void main() {
  print(Tool().pad('v', 3));
}"#,
        ["v3"]
    };

    mixin_multiple_with_preserves_all_getters => {
        r#"mixin Ga {
  int get a {
    return 1;
  }
}
mixin Gb {
  int get b {
    return 2;
  }
}
class Both with Ga, Gb {}
void main() {
  var b = Both();
  print(b.a + b.b);
}"#,
        ["3"]
    };
}
