//! Enhanced enums: fields, methods, const constructors, interface
//! implementation, and exhaustive switch patterns.

dart_cases! {
    enhanced_enum_single_int_field => {
        r#"enum Status {
  ok(200),
  fail(500);
  final int code;
  const Status(this.code);
}
void main() {
  print(Status.ok.code);
}"#,
        ["200"]
    };

    enhanced_enum_two_int_fields_sum => {
        r#"enum Point {
  origin(0, 0),
  unit(1, 1);
  final int x;
  final int y;
  const Point(this.x, this.y);
}
void main() {
  print(Point.unit.x + Point.unit.y);
}"#,
        ["2"]
    };

    enhanced_enum_string_field => {
        r#"enum Locale {
  en('English'),
  fr('French');
  final String label;
  const Locale(this.label);
}
void main() {
  print(Locale.fr.label);
}"#,
        ["French"]
    };

    enhanced_enum_bool_field => {
        r#"enum Switch {
  on(true),
  off(false);
  final bool active;
  const Switch(this.active);
}
void main() {
  print(Switch.on.active);
}"#,
        ["true"]
    };

    enhanced_enum_double_field => {
        r#"enum Rate {
  half(0.5),
  full(1.0);
  final double factor;
  const Rate(this.factor);
}
void main() {
  print(Rate.half.factor + Rate.full.factor);
}"#,
        ["1.5"]
    };

    enhanced_enum_method_returns_field => {
        r#"enum Http {
  get(1),
  post(2);
  final int id;
  const Http(this.id);
  int value() {
    return id;
  }
}
void main() {
  print(Http.post.value());
}"#,
        ["2"]
    };

    enhanced_enum_method_string_format => {
        r#"enum Color {
  red,
  blue;
  String hex() {
    if (this == Color.red) {
      return '#f00';
    }
    return '#00f';
  }
}
void main() {
  print(Color.red.hex());
}"#,
        ["#f00"]
    };

    enhanced_enum_getter_is_first => {
        r#"enum Rank {
  gold(1),
  silver(2),
  bronze(3);
  final int place;
  const Rank(this.place);
  bool get isWinner => place == 1;
}
void main() {
  print(Rank.gold.isWinner);
}"#,
        ["true"]
    };

    enhanced_enum_getter_not_winner => {
        r#"enum Rank {
  gold(1),
  silver(2);
  final int place;
  const Rank(this.place);
  bool get isWinner => place == 1;
}
void main() {
  print(Rank.silver.isWinner);
}"#,
        ["false"]
    };

    enhanced_enum_implements_interface => {
        r#"abstract class Describable {
  String describe();
}
enum Mood implements Describable {
  happy,
  sad;
  String describe() {
    return name;
  }
}
void main() {
  print(Mood.happy.describe());
}"#,
        ["happy"]
    };

    enhanced_enum_implements_with_field => {
        r#"abstract class Coded {
  int code();
}
enum ErrorCode implements Coded {
  notFound(404),
  server(500);
  final int value;
  const ErrorCode(this.value);
  int code() {
    return value;
  }
}
void main() {
  print(ErrorCode.notFound.code());
}"#,
        ["404"]
    };

    enhanced_enum_switch_exhaustive_two => {
        r#"enum Bin { zero, one }
String label(Bin b) {
  switch (b) {
    case Bin.zero:
      return '0';
    case Bin.one:
      return '1';
  }
}
void main() {
  print(label(Bin.one));
}"#,
        ["1"]
    };

    enhanced_enum_switch_exhaustive_three => {
        r#"enum Traffic { red, yellow, green }
int priority(Traffic t) {
  switch (t) {
    case Traffic.red:
      return 3;
    case Traffic.yellow:
      return 2;
    case Traffic.green:
      return 1;
  }
}
void main() {
  print(priority(Traffic.yellow));
}"#,
        ["2"]
    };

    enhanced_enum_switch_with_field_access => {
        r#"enum Level {
  low(1),
  high(10);
  final int power;
  const Level(this.power);
}
int boost(Level l) {
  switch (l) {
    case Level.low:
      return l.power;
    case Level.high:
      return l.power * 2;
  }
}
void main() {
  print(boost(Level.high));
}"#,
        ["20"]
    };

    enhanced_enum_const_constructor_three_args => {
        r#"enum Planet {
  earth(5.97e24, 6371, 1),
  mars(6.39e23, 3389, 2);
  final double mass;
  final double radius;
  final int order;
  const Planet(this.mass, this.radius, this.order);
}
void main() {
  print(Planet.mars.order);
}"#,
        ["2"]
    };

    enhanced_enum_method_compare_members => {
        r#"enum Size { small, medium, large }
bool isSmall(Size s) {
  return s == Size.small;
}
void main() {
  print(isSmall(Size.large));
}"#,
        ["false"]
    };

    enhanced_enum_method_on_weekday => {
        r#"enum Day {
  mon, tue, wed, thu, fri, sat, sun;
  bool isWeekend() {
    return this == Day.sat || this == Day.sun;
  }
}
void main() {
  print(Day.fri.isWeekend());
}"#,
        ["false"]
    };

    enhanced_enum_method_on_weekend => {
        r#"enum Day {
  mon, tue, wed, thu, fri, sat, sun;
  bool isWeekend() {
    return this == Day.sat || this == Day.sun;
  }
}
void main() {
  print(Day.sun.isWeekend());
}"#,
        ["true"]
    };

    enhanced_enum_toString_override => {
        r#"enum Token {
  alpha,
  beta;
  String display() {
    return 'Token.$name';
  }
}
void main() {
  print(Token.alpha.display());
}"#,
        ["Token.alpha"]
    };

    enhanced_enum_field_arithmetic => {
        r#"enum Op {
  add(1),
  mul(3);
  final int factor;
  const Op(this.factor);
}
void main() {
  print(Op.add.factor + Op.mul.factor);
}"#,
        ["4"]
    };

    enhanced_enum_list_field_length => {
        r#"enum Pack {
  small([1, 2]),
  big([1, 2, 3, 4]);
  final List<int> items;
  const Pack(this.items);
}
void main() {
  print(Pack.big.items.length);
}"#,
        ["4"]
    };

    enhanced_enum_map_field_lookup => {
        r#"enum Config {
  dev({'port': 3000}),
  prod({'port': 80});
  final Map<String, int> settings;
  const Config(this.settings);
}
void main() {
  print(Config.dev.settings['port']);
}"#,
        ["3000"]
    };

    enhanced_enum_switch_return_string => {
        r#"enum Dir { up, down, left, right }
String arrow(Dir d) {
  switch (d) {
    case Dir.up:
      return '^';
    case Dir.down:
      return 'v';
    case Dir.left:
      return '<';
    case Dir.right:
      return '>';
  }
}
void main() {
  print(arrow(Dir.left));
}"#,
        ["<"]
    };

    enhanced_enum_implements_two_methods => {
        r#"abstract class Named {
  String get label;
}
abstract class Valued {
  int get value;
}
enum Item implements Named, Valued {
  a(1),
  b(2);
  final int v;
  const Item(this.v);
  String get label => name;
  int get value => v;
}
void main() {
  print('${Item.b.label}:${Item.b.value}');
}"#,
        ["b:2"]
    };

    enhanced_enum_method_uses_index => {
        r#"enum Step {
  a, b, c;
  int position() {
    return index;
  }
}
void main() {
  print(Step.c.position());
}"#,
        ["2"]
    };

    enhanced_enum_field_negative => {
        r#"enum Sign {
  neg(-1),
  pos(1);
  final int val;
  const Sign(this.val);
}
void main() {
  print(Sign.neg.val + Sign.pos.val);
}"#,
        ["0"]
    };

    enhanced_enum_getter_from_field => {
        r#"enum Tier {
  free(0),
  pro(99);
  final int price;
  const Tier(this.price);
  bool get isFree => price == 0;
}
void main() {
  print(Tier.pro.isFree);
}"#,
        ["false"]
    };

    enhanced_enum_method_chain => {
        r#"enum Mode {
  read,
  write;
  String tag() {
    return name;
  }
  String full() {
    return 'mode:${tag()}';
  }
}
void main() {
  print(Mode.write.full());
}"#,
        ["mode:write"]
    };

    enhanced_enum_switch_in_main => {
        r#"enum State { idle, busy }
void main() {
  var s = State.busy;
  switch (s) {
    case State.idle:
      print('wait');
      break;
    case State.busy:
      print('work');
      break;
  }
}"#,
        ["work"]
    };

    enhanced_enum_four_members_field => {
        r#"enum Season {
  spring(1),
  summer(2),
  autumn(3),
  winter(4);
  final int month;
  const Season(this.month);
}
void main() {
  print(Season.winter.month);
}"#,
        ["4"]
    };

    enhanced_enum_method_param => {
        r#"enum MathOp {
  add,
  sub;
  int apply(int a, int b) {
    if (this == MathOp.add) {
      return a + b;
    }
    return a - b;
  }
}
void main() {
  print(MathOp.add.apply(10, 3));
}"#,
        ["13"]
    };

    enhanced_enum_subtraction_method => {
        r#"enum MathOp {
  add,
  sub;
  int apply(int a, int b) {
    if (this == MathOp.add) {
      return a + b;
    }
    return a - b;
  }
}
void main() {
  print(MathOp.sub.apply(10, 3));
}"#,
        ["7"]
    };

    enhanced_enum_static_like_method => {
        r#"enum Parse {
  intVal,
  strVal;
  String parseValue(Object v) {
    return v.toString();
  }
}
void main() {
  print(Parse.intVal.parseValue(42));
}"#,
        ["42"]
    };

    enhanced_enum_field_zero => {
        r#"enum Count {
  zero(0),
  one(1);
  final int n;
  const Count(this.n);
}
void main() {
  print(Count.zero.n);
}"#,
        ["0"]
    };

    enhanced_enum_field_large => {
        r#"enum Big {
  mega(1000000);
  final int size;
  const Big(this.size);
}
void main() {
  print(Big.mega.size);
}"#,
        ["1000000"]
    };

    enhanced_enum_switch_first_member => {
        r#"enum ABC { a, b, c }
String pick(ABC v) {
  switch (v) {
    case ABC.a:
      return 'first';
    case ABC.b:
      return 'mid';
    case ABC.c:
      return 'last';
  }
}
void main() {
  print(pick(ABC.a));
}"#,
        ["first"]
    };

    enhanced_enum_switch_last_member => {
        r#"enum ABC { a, b, c }
String pick(ABC v) {
  switch (v) {
    case ABC.a:
      return 'first';
    case ABC.b:
      return 'mid';
    case ABC.c:
      return 'last';
  }
}
void main() {
  print(pick(ABC.c));
}"#,
        ["last"]
    };

    enhanced_enum_interface_polymorphic => {
        r#"abstract class Describable {
  String describe();
}
enum Tag implements Describable {
  hot,
  cold;
  String describe() {
    return 'tag:$name';
  }
}
void main() {
  Describable d = Tag.hot;
  print(d.describe());
}"#,
        ["tag:hot"]
    };

    enhanced_enum_method_equality => {
        r#"enum Pair {
  aa,
  bb;
  bool matches(Pair other) {
    return this == other;
  }
}
void main() {
  print(Pair.aa.matches(Pair.aa));
}"#,
        ["true"]
    };

    enhanced_enum_method_inequality => {
        r#"enum Pair {
  aa,
  bb;
  bool matches(Pair other) {
    return this == other;
  }
}
void main() {
  print(Pair.aa.matches(Pair.bb));
}"#,
        ["false"]
    };

    enhanced_enum_field_string_concat => {
        r#"enum Greeting {
  hello('Hello'),
  bye('Goodbye');
  final String text;
  const Greeting(this.text);
}
void main() {
  print(Greeting.hello.text + ' World');
}"#,
        ["Hello World"]
    };

    enhanced_enum_getter_double => {
        r#"enum Ratio {
  half(0.5),
  quarter(0.25);
  final double value;
  const Ratio(this.value);
  double get doubled => value * 2;
}
void main() {
  print(Ratio.quarter.doubled);
}"#,
        ["0.5"]
    };

    enhanced_enum_method_in_values_loop => {
        r#"enum Digit {
  d0(0),
  d1(1),
  d2(2);
  final int num;
  const Digit(this.num);
}
void main() {
  var sum = 0;
  for (var d in Digit.values) {
    sum += d.num;
  }
  print(sum);
}"#,
        ["3"]
    };

    enhanced_enum_switch_with_method_call => {
        r#"enum Action {
  start,
  stop;
  String verb() {
    return name;
  }
}
String run(Action a) {
  switch (a) {
    case Action.start:
      return a.verb() + 'ed';
    case Action.stop:
      return a.verb() + 'ped';
  }
}
void main() {
  print(run(Action.start));
}"#,
        ["started"]
    };

    enhanced_enum_three_field_types => {
        r#"enum Record {
  entry(1, 'x', true);
  final int id;
  final String key;
  final bool active;
  const Record(this.id, this.key, this.active);
}
void main() {
  print('${Record.entry.id}:${Record.entry.key}:${Record.entry.active}');
}"#,
        ["1:x:true"]
    };

    enhanced_enum_method_returns_bool => {
        r#"enum Gate {
  open,
  closed;
  bool allows() {
    return this == Gate.open;
  }
}
void main() {
  print(Gate.closed.allows());
}"#,
        ["false"]
    };

    enhanced_enum_field_from_values => {
        r#"enum Code {
  a(10),
  b(20),
  c(30);
  final int val;
  const Code(this.val);
}
void main() {
  print(Code.values[1].val);
}"#,
        ["20"]
    };

    enhanced_enum_implements_describe_all => {
        r#"abstract class Describable {
  String describe();
}
enum Phase implements Describable {
  alpha,
  beta,
  gamma;
  String describe() {
    return 'phase-$name';
  }
}
void main() {
  print(Phase.beta.describe());
}"#,
        ["phase-beta"]
    };

    enhanced_enum_switch_five_members => {
        r#"enum Weekday { mon, tue, wed, thu, fri }
int num(Weekday w) {
  switch (w) {
    case Weekday.mon:
      return 1;
    case Weekday.tue:
      return 2;
    case Weekday.wed:
      return 3;
    case Weekday.thu:
      return 4;
    case Weekday.fri:
      return 5;
  }
}
void main() {
  print(num(Weekday.wed));
}"#,
        ["3"]
    };

    enhanced_enum_const_with_body_logic => {
        r#"enum Version {
  v1(1),
  v2(2);
  final int major;
  const Version(this.major);
  int next() {
    return major + 1;
  }
}
void main() {
  print(Version.v1.next());
}"#,
        ["2"]
    };

    enhanced_enum_nullable_field => {
        r#"enum Maybe {
  some(42),
  none(null);
  final int? value;
  const Maybe(this.value);
}
void main() {
  print(Maybe.none.value == null);
}"#,
        ["true"]
    };

    enhanced_enum_field_compare_in_method => {
        r#"enum Priority {
  low(1),
  high(5);
  final int weight;
  const Priority(this.weight);
  bool beats(Priority other) {
    return weight > other.weight;
  }
}
void main() {
  print(Priority.high.beats(Priority.low));
}"#,
        ["true"]
    };

    enhanced_enum_name_and_field => {
        r#"enum Fruit {
  apple(1),
  banana(2);
  final int count;
  const Fruit(this.count);
}
void main() {
  var f = Fruit.banana;
  print('${f.name}:${f.count}');
}"#,
        ["banana:2"]
    };

    enhanced_enum_switch_assign_var => {
        r#"enum Flag { on, off }
void main() {
  var f = Flag.on;
  var result = '';
  switch (f) {
    case Flag.on:
      result = 'yes';
      break;
    case Flag.off:
      result = 'no';
      break;
  }
  print(result);
}"#,
        ["yes"]
    };

    enhanced_enum_method_length_of_name => {
        r#"enum Key {
  ab,
  abcd;
  int nameLen() {
    return name.length;
  }
}
void main() {
  print(Key.abcd.nameLen());
}"#,
        ["4"]
    };

    enhanced_enum_two_member_implements => {
        r#"abstract class Printable {
  String printForm();
}
enum Status implements Printable {
  ok,
  err;
  String printForm() {
    return '[$name]';
  }
}
void main() {
  print(Status.err.printForm());
}"#,
        ["[err]"]
    };
}
