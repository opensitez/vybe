//! covariant on overridden method parameters and fields in inheritance.

dart_cases! {
    covariant_param_narrow_animal_to_cat => {
        r#"class Animal {}
class Cat extends Animal {}
class Cage {
  void admit(Animal a) {}
}
class CatCage extends Cage {
  @override
  void admit(covariant Cat c) {}
}
void main() {
  var cage = CatCage();
  cage.admit(Cat());
  print('ok');
}"#,
        ["ok"]
    };

    covariant_param_narrow_num_to_int => {
        r#"class NumHolder {
  void store(num n) {}
}
class IntHolder extends NumHolder {
  @override
  void store(covariant int n) {}
}
void main() {
  IntHolder().store(42);
  print(42);
}"#,
        ["42"]
    };

    covariant_param_narrow_object_to_string => {
        r#"class Box {
  void pack(Object o) {}
}
class StringBox extends Box {
  @override
  void pack(covariant String s) {}
}
void main() {
  StringBox().pack('data');
  print('data');
}"#,
        ["data"]
    };

    covariant_param_used_in_override_body => {
        r#"class Animal {
  String name;
  Animal(this.name);
}
class Dog extends Animal {
  Dog(String n) : super(n);
}
class Trainer {
  void train(Animal a) {}
}
class DogTrainer extends Trainer {
  @override
  void train(covariant Dog d) {
    print(d.name);
  }
}
void main() {
  DogTrainer().train(Dog('Rex'));
}"#,
        ["Rex"]
    };

    covariant_param_two_level_hierarchy => {
        r#"class Vehicle {}
class Car extends Vehicle {}
class Garage {
  void park(Vehicle v) {}
}
class CarGarage extends Garage {
  @override
  void park(covariant Car c) {}
}
void main() {
  CarGarage().park(Car());
  print(1);
}"#,
        ["1"]
    };

    covariant_param_with_return_value => {
        r#"class Node {}
class Leaf extends Node {}
class Tree {
  String tag(Node n) {
    return 'node';
  }
}
class LeafTree extends Tree {
  @override
  String tag(covariant Leaf l) {
    return 'leaf';
  }
}
void main() {
  print(LeafTree().tag(Leaf()));
}"#,
        ["leaf"]
    };

    covariant_param_bool_subclass => {
        r#"class Flag {
  bool on;
  Flag(this.on);
}
class Switch extends Flag {
  Switch(bool v) : super(v);
}
class Panel {
  void flip(Flag f) {}
}
class SwitchPanel extends Panel {
  @override
  void flip(covariant Switch s) {
    print(s.on);
  }
}
void main() {
  SwitchPanel().flip(Switch(true));
}"#,
        ["true"]
    };

    covariant_param_list_element_type => {
        r#"class Container {
  void add(List<Object> items) {}
}
class IntContainer extends Container {
  @override
  void add(covariant List<int> items) {
    print(items.length);
  }
}
void main() {
  IntContainer().add([1, 2, 3]);
}"#,
        ["3"]
    };

    covariant_param_map_value_type => {
        r#"class Registry {
  void register(Map<String, Object> m) {}
}
class IntRegistry extends Registry {
  @override
  void register(covariant Map<String, int> m) {
    print(m['x']);
  }
}
void main() {
  IntRegistry().register({'x': 9});
}"#,
        ["9"]
    };

    covariant_field_int_over_num_getter => {
        r#"class NumBox {
  num get value => 0;
}
class IntBox extends NumBox {
  @override
  covariant int value = 7;
}
void main() {
  print(IntBox().value);
}"#,
        ["7"]
    };

    covariant_field_string_over_object => {
        r#"class ObjSlot {
  Object get slot => 'x';
}
class StrSlot extends ObjSlot {
  @override
  covariant String slot = 'hello';
}
void main() {
  print(StrSlot().slot);
}"#,
        ["hello"]
    };

    covariant_field_mutable_in_subclass => {
        r#"class Base {
  num get n => 1;
}
class Sub extends Base {
  @override
  covariant int n = 5;
}
void main() {
  var s = Sub();
  s.n = 10;
  print(s.n);
}"#,
        ["10"]
    };

    covariant_param_three_overrides_in_chain => {
        r#"class A {}
class B extends A {}
class C extends B {}
class Handler {
  void handle(A a) {}
}
class BHandler extends Handler {
  @override
  void handle(covariant B b) {}
}
class CHandler extends BHandler {
  @override
  void handle(covariant C c) {
    print('c');
  }
}
void main() {
  CHandler().handle(C());
}"#,
        ["c"]
    };

    covariant_param_with_super_call => {
        r#"class Base {
  int n;
  Base(this.n);
}
class Derived extends Base {
  Derived(int v) : super(v);
}
class Processor {
  void run(Base b) {}
}
class DerivedProcessor extends Processor {
  @override
  void run(covariant Derived d) {
    print(d.n);
  }
}
void main() {
  DerivedProcessor().run(Derived(6));
}"#,
        ["6"]
    };

    covariant_param_nullable_narrowing => {
        r#"class Maybe {
  void set(Object? v) {}
}
class IntMaybe extends Maybe {
  @override
  void set(covariant int? v) {
    print(v ?? 0);
  }
}
void main() {
  IntMaybe().set(3);
}"#,
        ["3"]
    };

    covariant_param_nullable_sets_null => {
        r#"class Maybe {
  void set(Object? v) {}
}
class IntMaybe extends Maybe {
  @override
  void set(covariant int? v) {
    print(v == null);
  }
}
void main() {
  IntMaybe().set(null);
}"#,
        ["true"]
    };

    covariant_field_list_of_int => {
        r#"class AnyList {
  List<Object> get items => [];
}
class IntList extends AnyList {
  @override
  covariant List<int> items = [1, 2];
}
void main() {
  print(IntList().items[1]);
}"#,
        ["2"]
    };

    covariant_param_method_count_items => {
        r#"class Collection {
  int count(List<Object> xs) {
    return xs.length;
  }
}
class IntCollection extends Collection {
  @override
  int count(covariant List<int> xs) {
    return xs.length + 1;
  }
}
void main() {
  print(IntCollection().count([1, 2]));
}"#,
        ["3"]
    };

    covariant_param_shape_hierarchy => {
        r#"class Shape {}
class Circle extends Shape {
  int r;
  Circle(this.r);
}
class Drawer {
  void draw(Shape s) {}
}
class CircleDrawer extends Drawer {
  @override
  void draw(covariant Circle c) {
    print(c.r);
  }
}
void main() {
  CircleDrawer().draw(Circle(5));
}"#,
        ["5"]
    };

    covariant_param_employee_manager => {
        r#"class Person {
  String name;
  Person(this.name);
}
class Employee extends Person {
  Employee(String n) : super(n);
}
class Dept {
  void hire(Person p) {}
}
class HR extends Dept {
  @override
  void hire(covariant Employee e) {
    print(e.name);
  }
}
void main() {
  HR().hire(Employee('Sam'));
}"#,
        ["Sam"]
    };

    covariant_field_double_over_num => {
        r#"class NumVal {
  num get reading => 0.0;
}
class DoubleVal extends NumVal {
  @override
  covariant double reading = 3.14;
}
void main() {
  print(DoubleVal().reading > 3.0);
}"#,
        ["true"]
    };

    covariant_param_event_handler => {
        r#"class Event {}
class Click extends Event {
  int x;
  Click(this.x);
}
class Listener {
  void on(Event e) {}
}
class ClickListener extends Listener {
  @override
  void on(covariant Click e) {
    print(e.x);
  }
}
void main() {
  ClickListener().on(Click(99));
}"#,
        ["99"]
    };

    covariant_param_compare_type_tag => {
        r#"class Item {}
class Book extends Item {
  String title;
  Book(this.title);
}
class Shelf {
  String label(Item i) {
    return 'item';
  }
}
class BookShelf extends Shelf {
  @override
  String label(covariant Book b) {
    return b.title;
  }
}
void main() {
  print(BookShelf().label(Book('Dart')));
}"#,
        ["Dart"]
    };

    covariant_param_food_chain => {
        r#"class Food {}
class Fruit extends Food {}
class Apple extends Fruit {}
class Eater {
  void eat(Food f) {}
}
class FruitEater extends Eater {
  @override
  void eat(covariant Fruit f) {}
}
class AppleEater extends FruitEater {
  @override
  void eat(covariant Apple a) {
    print('apple');
  }
}
void main() {
  AppleEater().eat(Apple());
}"#,
        ["apple"]
    };

    covariant_field_bool_flag => {
        r#"class Toggle {
  Object get state => false;
}
class BoolToggle extends Toggle {
  @override
  covariant bool state = true;
}
void main() {
  print(BoolToggle().state);
}"#,
        ["true"]
    };

    covariant_param_writer_narrow => {
        r#"class Writer {
  void write(Object data) {}
}
class TextWriter extends Writer {
  @override
  void write(covariant String data) {
    print(data.length);
  }
}
void main() {
  TextWriter().write('abc');
}"#,
        ["3"]
    };

    covariant_param_reader_returns_length => {
        r#"class Reader {
  int size(List<Object> buf) {
    return buf.length;
  }
}
class ByteReader extends Reader {
  @override
  int size(covariant List<int> buf) {
    return buf.length * 2;
  }
}
void main() {
  print(ByteReader().size([1, 2, 3, 4]));
}"#,
        ["8"]
    };

    covariant_param_account_hierarchy => {
        r#"class Account {
  int balance;
  Account(this.balance);
}
class Savings extends Account {
  Savings(int b) : super(b);
}
class Bank {
  void deposit(Account a, int amt) {}
}
class SavingsBank extends Bank {
  @override
  void deposit(covariant Savings s, int amt) {
    print(s.balance + amt);
  }
}
void main() {
  SavingsBank().deposit(Savings(100), 50);
}"#,
        ["150"]
    };

    covariant_field_map_string_int => {
        r#"class AnyMap {
  Map<Object, Object> get data => {};
}
class StrIntMap extends AnyMap {
  @override
  covariant Map<String, int> data = {'k': 1};
}
void main() {
  print(StrIntMap().data['k']);
}"#,
        ["1"]
    };

    covariant_param_media_play => {
        r#"class Media {}
class Audio extends Media {
  int duration;
  Audio(this.duration);
}
class Player {
  void play(Media m) {}
}
class AudioPlayer extends Player {
  @override
  void play(covariant Audio a) {
    print(a.duration);
  }
}
void main() {
  AudioPlayer().play(Audio(120));
}"#,
        ["120"]
    };

    covariant_param_error_type_narrow => {
        r#"class ErrorBase {}
class NetworkError extends ErrorBase {
  int code;
  NetworkError(this.code);
}
class Handler {
  void handle(ErrorBase e) {}
}
class NetHandler extends Handler {
  @override
  void handle(covariant NetworkError e) {
    print(e.code);
  }
}
void main() {
  NetHandler().handle(NetworkError(404));
}"#,
        ["404"]
    };

    covariant_param_geometry_point => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
}
class ColoredPoint extends Point {
  String color;
  ColoredPoint(int x, int y, this.color) : super(x, y);
}
class Plotter {
  void mark(Point p) {}
}
class ColorPlotter extends Plotter {
  @override
  void mark(covariant ColoredPoint p) {
    print('${p.x},${p.color}');
  }
}
void main() {
  ColorPlotter().mark(ColoredPoint(1, 2, 'red'));
}"#,
        ["1,red"]
    };

    covariant_field_zero_initial => {
        r#"class Counter {
  num get count => 0;
}
class IntCounter extends Counter {
  @override
  covariant int count = 0;
}
void main() {
  print(IntCounter().count);
}"#,
        ["0"]
    };

    covariant_param_logger_message => {
        r#"class Logger {
  void log(Object msg) {}
}
class StringLogger extends Logger {
  @override
  void log(covariant String msg) {
    print(msg.toUpperCase());
  }
}
void main() {
  StringLogger().log('hi');
}"#,
        ["HI"]
    };

    covariant_param_stack_push => {
        r#"class Stack {
  void push(Object item) {}
}
class IntStack extends Stack {
  int total = 0;
  @override
  void push(covariant int item) {
    total = total + item;
  }
}
void main() {
  var s = IntStack();
  s.push(3);
  s.push(4);
  print(s.total);
}"#,
        ["7"]
    };

    covariant_param_animal_shelter_adopt => {
        r#"class Pet {
  String species;
  Pet(this.species);
}
class Cat extends Pet {
  Cat() : super('cat');
}
class Shelter {
  String adopt(Pet p) {
    return p.species;
  }
}
class CatShelter extends Shelter {
  @override
  String adopt(covariant Cat c) {
    return 'kitten';
  }
}
void main() {
  print(CatShelter().adopt(Cat()));
}"#,
        ["kitten"]
    };

    covariant_field_negative_int => {
        r#"class Signed {
  num get val => 0;
}
class Negative extends Signed {
  @override
  covariant int val = -5;
}
void main() {
  print(Negative().val);
}"#,
        ["-5"]
    };

    covariant_param_config_merge => {
        r#"class Config {
  void merge(Map<String, Object> opts) {}
}
class AppConfig extends Config {
  int port = 80;
  @override
  void merge(covariant Map<String, int> opts) {
    if (opts.containsKey('port')) {
      port = opts['port']!;
    }
  }
}
void main() {
  var c = AppConfig();
  c.merge({'port': 3000});
  print(c.port);
}"#,
        ["3000"]
    };

    covariant_param_document_print => {
        r#"class Document {}
class Pdf extends Document {
  int pages;
  Pdf(this.pages);
}
class Printer {
  void printDoc(Document d) {}
}
class PdfPrinter extends Printer {
  @override
  void printDoc(covariant Pdf d) {
    print(d.pages);
  }
}
void main() {
  PdfPrinter().printDoc(Pdf(10));
}"#,
        ["10"]
    };

    covariant_field_large_string => {
        r#"class Holder {
  Object get payload => '';
}
class TextHolder extends Holder {
  @override
  covariant String payload = 'longtext';
}
void main() {
  print(TextHolder().payload.length);
}"#,
        ["8"]
    };
}
