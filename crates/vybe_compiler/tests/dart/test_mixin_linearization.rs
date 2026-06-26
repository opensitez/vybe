//! Mixin linearization: declaration order, super resolution,
//! on Type constraints, and conflict precedence.

dart_cases! {
    mixin_order_ab_resolves_b_method => {
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
class X with A, B {}
void main() {
  print(X().tag());
}"#,
        ["b"]
    };

    mixin_order_ba_resolves_a_method => {
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
class Y with B, A {}
void main() {
  print(Y().tag());
}"#,
        ["a"]
    };

    three_mixins_last_declared_wins_conflict => {
        r#"mixin M1 {
  int val() {
    return 1;
  }
}
mixin M2 {
  int val() {
    return 2;
  }
}
mixin M3 {
  int val() {
    return 3;
  }
}
class Z with M1, M2, M3 {}
void main() {
  print(Z().val());
}"#,
        ["3"]
    };

    reversed_three_mixins_first_wins => {
        r#"mixin M1 {
  int val() {
    return 1;
  }
}
mixin M2 {
  int val() {
    return 2;
  }
}
mixin M3 {
  int val() {
    return 3;
  }
}
class W with M3, M2, M1 {}
void main() {
  print(W().val());
}"#,
        ["1"]
    };

    super_in_mixin_calls_next_mixin_implementation => {
        r#"mixin A {
  String run() {
    return 'A';
  }
}
mixin B on Object {
  String run() {
    return super.run() + 'B';
  }
}
class C with A, B {}
void main() {
  print(C().run());
}"#,
        ["AB"]
    };

    super_in_first_mixin_reaches_class_supertype => {
        r#"class Base {
  String root() {
    return 'base';
  }
}
mixin Wrap on Base {
  String root() {
    return super.root() + '-wrap';
  }
}
class Node extends Base with Wrap {}
void main() {
  print(Node().root());
}"#,
        ["base-wrap"]
    };

    mixin_on_constraint_allows_super_call => {
        r#"class Engine {
  int power = 10;
}
mixin Turbo on Engine {
  int boosted() {
    return super.power + 5;
  }
}
class Car extends Engine with Turbo {}
void main() {
  print(Car().boosted());
}"#,
        ["15"]
    };

    mixin_on_reads_superclass_field_directly => {
        r#"class Vehicle {
  int wheels = 4;
}
mixin Inspect on Vehicle {
  int countWheels() {
    return wheels;
  }
}
class Truck extends Vehicle with Inspect {}
void main() {
  print(Truck().countWheels());
}"#,
        ["4"]
    };

    mixin_on_mutates_superclass_field => {
        r#"class Counter {
  int n = 0;
}
mixin Bump on Counter {
  void inc() {
    n = n + 1;
  }
}
class Tally extends Counter with Bump {}
void main() {
  var t = Tally();
  t.inc();
  print(t.n);
}"#,
        ["1"]
    };

    mixin_on_abstract_supertype_implementation => {
        r#"abstract class Shape {
  int sides();
}
mixin Polygon on Shape {
  String kind() {
    return 'poly';
  }
}
class Tri extends Shape with Polygon {
  int sides() {
    return 3;
  }
}
void main() {
  print(Tri().sides());
}"#,
        ["3"]
    };

    extends_then_with_mixin_order => {
        r#"class Base {
  String from() {
    return 'base';
  }
}
mixin Mid {
  String from() {
    return 'mid';
  }
}
class Leaf extends Base with Mid {}
void main() {
  print(Leaf().from());
}"#,
        ["mid"]
    };

    class_method_overridden_by_rightmost_mixin => {
        r#"class Host {
  int score() {
    return 1;
  }
}
mixin Boost {
  int score() {
    return 10;
  }
}
class Player extends Host with Boost {}
void main() {
  print(Player().score());
}"#,
        ["10"]
    };

    left_mixin_non_conflicting_methods_both_available => {
        r#"mixin Left {
  int leftVal() {
    return 1;
  }
}
mixin Right {
  int rightVal() {
    return 2;
  }
}
class Both with Left, Right {}
void main() {
  print(Both().leftVal() + Both().rightVal());
}"#,
        ["3"]
    };

    mixin_super_chain_three_levels => {
        r#"mixin A {
  String chain() {
    return 'A';
  }
}
mixin B on Object {
  String chain() {
    return super.chain() + 'B';
  }
}
mixin C on Object {
  String chain() {
    return super.chain() + 'C';
  }
}
class D with A, B, C {}
void main() {
  print(D().chain());
}"#,
        ["ABC"]
    };

    mixin_on_requires_extends_supertype => {
        r#"class Animal {
  String name = 'pet';
}
mixin Named on Animal {
  String label() {
    return name;
  }
}
class Dog extends Animal with Named {}
void main() {
  print(Dog().label());
}"#,
        ["pet"]
    };

    two_mixins_same_name_different_return_paths => {
        r#"mixin Fast {
  String mode() {
    return 'fast';
  }
}
mixin Slow {
  String mode() {
    return 'slow';
  }
}
class Runner with Fast, Slow {}
void main() {
  print(Runner().mode());
}"#,
        ["slow"]
    };

    swap_mixin_order_changes_winner => {
        r#"mixin First {
  int pick() {
    return 100;
  }
}
mixin Second {
  int pick() {
    return 200;
  }
}
class One with First, Second {}
class Two with Second, First {}
void main() {
  print(One().pick() + Two().pick());
}"#,
        ["300"]
    };

    mixin_method_calls_superclass_method_via_on => {
        r#"class Base {
  int compute() {
    return 5;
  }
}
mixin Double on Base {
  int compute() {
    return super.compute() * 2;
  }
}
class Twice extends Base with Double {}
void main() {
  print(Twice().compute());
}"#,
        ["10"]
    };

    mixin_on_intermediate_subclass => {
        r#"class Root {
  int depth = 1;
}
class Mid extends Root {}
mixin Tag on Mid {
  int tagged() {
    return depth + 10;
  }
}
class Leaf extends Mid with Tag {}
void main() {
  print(Leaf().tagged());
}"#,
        ["11"]
    };

    four_mixins_cascading_super_calls => {
        r#"mixin A {
  int n() {
    return 1;
  }
}
mixin B on Object {
  int n() {
    return super.n() + 2;
  }
}
mixin C on Object {
  int n() {
    return super.n() + 3;
  }
}
mixin D on Object {
  int n() {
    return super.n() + 4;
  }
}
class E with A, B, C, D {}
void main() {
  print(E().n());
}"#,
        ["10"]
    };

    mixin_field_not_shadowed_by_class_field => {
        r#"mixin M {
  int x = 100;
}
class C with M {
  int y = 1;
}
void main() {
  print(C().x + C().y);
}"#,
        ["101"]
    };

    mixin_getter_overrides_class_getter => {
        r#"class Base {
  int get val {
    return 1;
  }
}
mixin Override {
  int get val {
    return 9;
  }
}
class Sub extends Base with Override {}
void main() {
  print(Sub().val);
}"#,
        ["9"]
    };

    mixin_on_with_super_in_getter => {
        r#"class Base {
  int get base {
    return 2;
  }
}
mixin Extra on Base {
  int get total {
    return super.base + 3;
  }
}
class Node extends Base with Extra {}
void main() {
  print(Node().total);
}"#,
        ["5"]
    };

    diamond_like_mixin_resolution_last_wins => {
        r#"mixin X {
  String id() {
    return 'x';
  }
}
mixin Y {
  String id() {
    return 'y';
  }
}
class Combo with X, Y {}
void main() {
  print(Combo().id());
}"#,
        ["y"]
    };

    mixin_applied_to_class_without_extends => {
        r#"mixin Solo {
  int one() {
    return 1;
  }
}
class Plain with Solo {}
void main() {
  print(Plain().one());
}"#,
        ["1"]
    };

    super_in_mixin_on_object_reaches_implicit_super => {
        r#"mixin M on Object {
  String describe() {
    return 'mixin';
  }
}
class T with M {}
void main() {
  print(T().describe());
}"#,
        ["mixin"]
    };

    mixin_order_affects_only_conflicting_members => {
        r#"mixin Alpha {
  int a() {
    return 1;
  }
  int clash() {
    return 10;
  }
}
mixin Beta {
  int b() {
    return 2;
  }
  int clash() {
    return 20;
  }
}
class Gamma with Alpha, Beta {}
void main() {
  var g = Gamma();
  print(g.a() + g.b() + g.clash());
}"#,
        ["23"]
    };

    mixin_on_base_with_subclass_constructor => {
        r#"class Base {
  int v;
  Base(this.v);
}
mixin Scale on Base {
  int scaled() {
    return v * 2;
  }
}
class Sub extends Base with Scale {
  Sub(int x) : super(x);
}
void main() {
  print(Sub(4).scaled());
}"#,
        ["8"]
    };

    two_on_constraints_stacked_hierarchy => {
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

    mixin_super_invokes_parent_mixin_method => {
        r#"mixin P {
  String step() {
    return 'P';
  }
}
mixin Q on Object {
  String step() {
    return super.step() + 'Q';
  }
}
class R with P, Q {}
void main() {
  print(R().step());
}"#,
        ["PQ"]
    };

    mixin_method_uses_host_instance_state => {
        r#"mixin Stateful {
  int ticks = 0;
  void tick() {
    ticks = ticks + 1;
  }
}
class Clock with Stateful {}
void main() {
  var c = Clock();
  c.tick();
  c.tick();
  print(c.ticks);
}"#,
        ["2"]
    };

    rightmost_mixin_wins_for_operator_like_method => {
        r#"mixin AddOne {
  int transform(int n) {
    return n + 1;
  }
}
mixin AddTen {
  int transform(int n) {
    return n + 10;
  }
}
class Pipe with AddOne, AddTen {}
void main() {
  print(Pipe().transform(5));
}"#,
        ["15"]
    };

    mixin_on_supertype_method_visible_in_mixin => {
        r#"class Parent {
  String greet() {
    return 'hi';
  }
}
mixin Child on Parent {
  String shout() {
    return greet().toUpperCase();
  }
}
class Kid extends Parent with Child {}
void main() {
  print(Kid().shout());
}"#,
        ["HI"]
    };

    linearization_preserves_non_overridden_super_methods => {
        r#"class Base {
  int baseOnly() {
    return 7;
  }
}
mixin M {
  int mixOnly() {
    return 3;
  }
}
class Both extends Base with M {}
void main() {
  var b = Both();
  print(b.baseOnly() + b.mixOnly());
}"#,
        ["10"]
    };

    mixin_conflict_resolution_with_extends_and_with => {
        r#"class Root {
  String who() {
    return 'root';
  }
}
mixin Layer {
  String who() {
    return 'layer';
  }
}
class Leaf extends Root with Layer {}
void main() {
  print(Leaf().who());
}"#,
        ["layer"]
    };

    super_in_on_mixin_delegates_to_extended_class => {
        r#"class Base {
  int value = 3;
}
mixin Times on Base {
  int value = 99;
  int read() {
    return super.value;
  }
}
class Box extends Base with Times {}
void main() {
  print(Box().read());
}"#,
        ["3"]
    };

    five_mixins_rightmost_tag_wins => {
        r#"mixin T1 {
  String tag() {
    return '1';
  }
}
mixin T2 {
  String tag() {
    return '2';
  }
}
mixin T3 {
  String tag() {
    return '3';
  }
}
mixin T4 {
  String tag() {
    return '4';
  }
}
mixin T5 {
  String tag() {
    return '5';
  }
}
class All with T1, T2, T3, T4, T5 {}
void main() {
  print(All().tag());
}"#,
        ["5"]
    };

    mixin_on_with_named_constructor_on_subclass => {
        r#"class Base {
  int n;
  Base(this.n);
}
mixin Double on Base {
  int twice() {
    return n * 2;
  }
}
class Sub extends Base with Double {
  Sub.zero() : super(0);
}
void main() {
  print(Sub.zero().twice());
}"#,
        ["0"]
    };

    mixin_order_change_alters_super_chain_output => {
        r#"mixin A {
  String build() {
    return 'a';
  }
}
mixin B on Object {
  String build() {
    return super.build() + 'b';
  }
}
class AB with A, B {}
class BA with B, A {}
void main() {
  print(AB().build().length + BA().build().length);
}"#,
        ["4"]
    };

    mixin_on_abstract_with_concrete_subclass => {
        r#"abstract class Repo {
  int size();
}
mixin Cache on Repo {
  String label() {
    return 'cached';
  }
}
class ListRepo extends Repo with Cache {
  int size() {
    return 5;
  }
}
void main() {
  print(ListRepo().size());
}"#,
        ["5"]
    };

    multiple_mixins_each_add_distinct_getter => {
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
mixin Gc {
  int get c {
    return 3;
  }
}
class All with Ga, Gb, Gc {}
void main() {
  var x = All();
  print(x.a + x.b + x.c);
}"#,
        ["6"]
    };

    mixin_super_call_with_on_two_level_hierarchy => {
        r#"class Grand {
  int g() {
    return 1;
  }
}
class Parent extends Grand {}
mixin Child on Parent {
  int total() {
    return super.g() + 5;
  }
}
class Leaf extends Parent with Child {}
void main() {
  print(Leaf().total());
}"#,
        ["6"]
    };

    reversed_pair_mixins_opposite_winners => {
        r#"mixin Red {
  int hue() {
    return 1;
  }
}
mixin Blue {
  int hue() {
    return 2;
  }
}
class RB with Red, Blue {}
class BR with Blue, Red {}
void main() {
  print(RB().hue() + BR().hue());
}"#,
        ["3"]
    };

    mixin_on_constraint_method_composes_with_class_method => {
        r#"class Worker {
  int basePay() {
    return 100;
  }
}
mixin Bonus on Worker {
  int totalPay() {
    return basePay() + 50;
  }
}
class Employee extends Worker with Bonus {}
void main() {
  print(Employee().totalPay());
}"#,
        ["150"]
    };

    mixin_linearization_instance_per_class => {
        r#"mixin Tag {
  String name = 'tagged';
}
class A with Tag {}
class B with Tag {}
void main() {
  var a = A();
  var b = B();
  b.name = 'changed';
  print(a.name);
}"#,
        ["tagged"]
    };
}
