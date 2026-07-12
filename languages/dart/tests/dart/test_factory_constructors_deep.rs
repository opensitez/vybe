//! Factory constructors: singleton caches, redirects, validation,
//! fromJson patterns, and multi-path creation logic.

dart_cases! {
    factory_singleton_returns_same_instance => {
        r#"class Logger {
  static Logger? _inst;
  int hits = 0;
  Logger._();
  factory Logger() {
    _inst ??= Logger._();
    return _inst!;
  }
}
void main() {
  var a = Logger();
  var b = Logger();
  print(a == b);
}"#,
        ["true"]
    };

    factory_singleton_preserves_mutated_state => {
        r#"class Cache {
  static Cache? _one;
  int count = 0;
  Cache._();
  factory Cache() {
    _one ??= Cache._();
    return _one!;
  }
}
void main() {
  var a = Cache();
  a.count = 5;
  print(Cache().count);
}"#,
        ["5"]
    };

    factory_redirect_arrow_to_private_constructor => {
        r#"class Token {
  final String value;
  Token._(this.value);
  factory Token(String v) => Token._(v);
}
void main() {
  print(Token('secret').value);
}"#,
        ["secret"]
    };

    factory_redirect_equals_syntax => {
        r#"class Point {
  int x;
  int y;
  Point(this.x, this.y);
  Point._zero() : x = 0, y = 0;
  factory Point.zero() = Point._zero;
}
void main() {
  print(Point.zero().x);
}"#,
        ["0"]
    };

    factory_assert_validates_positive_input => {
        r#"class Age {
  int years;
  Age._(this.years);
  factory Age(int y) {
    assert(y >= 0);
    return Age._(y);
  }
}
void main() {
  print(Age(30).years);
}"#,
        ["30"]
    };

    factory_assert_with_message_pattern => {
        r#"class Port {
  int number;
  Port._(this.number);
  factory Port(int n) {
    assert(n > 0, 'port must be positive');
    return Port._(n);
  }
}
void main() {
  print(Port(443).number);
}"#,
        ["443"]
    };

    factory_from_json_parses_int_field => {
        r#"class User {
  int id;
  String name;
  User._(this.id, this.name);
  factory User.fromJson(Map<String, dynamic> json) {
    return User._(json['id'], json['name']);
  }
}
void main() {
  var u = User.fromJson({'id': 7, 'name': 'Ann'});
  print(u.id);
}"#,
        ["7"]
    };

    factory_from_json_reads_string_field => {
        r#"class User {
  int id;
  String name;
  User._(this.id, this.name);
  factory User.fromJson(Map<String, dynamic> json) {
    return User._(json['id'], json['name']);
  }
}
void main() {
  var u = User.fromJson({'id': 1, 'name': 'Bob'});
  print(u.name);
}"#,
        ["Bob"]
    };

    factory_from_json_with_default_for_missing_key => {
        r#"class Config {
  int timeout;
  Config._(this.timeout);
  factory Config.fromJson(Map<String, dynamic> json) {
    var t = json['timeout'];
    return Config._(t ?? 30);
  }
}
void main() {
  print(Config.fromJson({}).timeout);
}"#,
        ["30"]
    };

    factory_from_json_nested_map => {
        r#"class Address {
  String city;
  Address._(this.city);
  factory Address.fromJson(Map<String, dynamic> json) {
    return Address._(json['city']);
  }
}
class Person {
  String name;
  Address addr;
  Person._(this.name, this.addr);
  factory Person.fromJson(Map<String, dynamic> json) {
    return Person._(json['name'], Address.fromJson(json['addr']));
  }
}
void main() {
  var p = Person.fromJson({'name': 'Eve', 'addr': {'city': 'Paris'}});
  print(p.addr.city);
}"#,
        ["Paris"]
    };

    factory_cached_by_string_key => {
        r#"class Icon {
  static final Map<String, Icon> _cache = {};
  String name;
  Icon._(this.name);
  factory Icon(String n) {
    return _cache.putIfAbsent(n, () => Icon._(n));
  }
}
void main() {
  var a = Icon('home');
  var b = Icon('home');
  print(a == b);
}"#,
        ["true"]
    };

    factory_returns_different_instances_for_different_keys => {
        r#"class Icon {
  static final Map<String, Icon> _cache = {};
  String name;
  Icon._(this.name);
  factory Icon(String n) {
    return _cache.putIfAbsent(n, () => Icon._(n));
  }
}
void main() {
  print(Icon('a') == Icon('b'));
}"#,
        ["false"]
    };

    factory_selects_subclass_by_type_code => {
        r#"class Shape {
  int sides;
  Shape(this.sides);
  factory Shape.fromCode(String code) {
    if (code == 'tri') {
      return Triangle();
    }
    return Square();
  }
}
class Triangle extends Shape {
  Triangle() : super(3);
}
class Square extends Shape {
  Square() : super(4);
}
void main() {
  print(Shape.fromCode('tri').sides);
}"#,
        ["3"]
    };

    factory_named_zero_returns_constant_like_instance => {
        r#"class Vec {
  int x;
  int y;
  Vec(this.x, this.y);
  factory Vec.zero() {
    return Vec(0, 0);
  }
}
void main() {
  print(Vec.zero().x + Vec.zero().y);
}"#,
        ["0"]
    };

    factory_with_conditional_creation_path => {
        r#"class Num {
  int v;
  Num(this.v);
  factory Num.parse(String s) {
    if (s == 'zero') {
      return Num(0);
    }
    return Num(int.parse(s));
  }
}
void main() {
  print(Num.parse('12').v);
}"#,
        ["12"]
    };

    factory_private_constructor_with_validation => {
        r#"class Email {
  String addr;
  Email._(this.addr);
  factory Email(String raw) {
    assert(raw.contains('@'));
    return Email._(raw);
  }
}
void main() {
  print(Email('a@b.c').addr);
}"#,
        ["a@b.c"]
    };

    factory_redirect_named_to_generative_named => {
        r#"class Pair {
  int a;
  int b;
  Pair(this.a, this.b);
  Pair.same(int v) : a = v, b = v;
  factory Pair.fromSame(int v) => Pair.same(v);
}
void main() {
  var p = Pair.fromSame(3);
  print(p.a + p.b);
}"#,
        ["6"]
    };

    factory_registry_lookup_or_create => {
        r#"class Service {
  static final Map<int, Service> _registry = {};
  int id;
  Service._(this.id);
  factory Service.forId(int id) {
    return _registry.putIfAbsent(id, () => Service._(id));
  }
}
void main() {
  print(Service.forId(9).id);
}"#,
        ["9"]
    };

    factory_from_json_list_field_length => {
        r#"class Batch {
  List<int> ids;
  Batch._(this.ids);
  factory Batch.fromJson(Map<String, dynamic> json) {
    var raw = json['ids'] as List;
    return Batch._(raw.cast<int>());
  }
}
void main() {
  var b = Batch.fromJson({'ids': [1, 2, 3]});
  print(b.ids.length);
}"#,
        ["3"]
    };

    factory_flyweight_reuses_equal_value_objects => {
        r#"class SmallInt {
  static final Map<int, SmallInt> _pool = {};
  int value;
  SmallInt._(this.value);
  factory SmallInt(int v) {
    return _pool.putIfAbsent(v, () => SmallInt._(v));
  }
}
void main() {
  print(SmallInt(5) == SmallInt(5));
}"#,
        ["true"]
    };

    factory_returns_new_each_time_without_cache => {
        r#"class Box {
  int id;
  static int _next = 0;
  Box._(this.id);
  factory Box() {
    _next = _next + 1;
    return Box._(_next);
  }
}
void main() {
  print(Box().id + Box().id);
}"#,
        ["3"]
    };

    factory_with_static_counter_increments => {
        r#"class Ticket {
  static int serial = 0;
  int number;
  Ticket(this.number);
  factory Ticket.next() {
    serial = serial + 1;
    return Ticket(serial);
  }
}
void main() {
  print(Ticket.next().number);
  print(Ticket.next().number);
}"#,
        ["1", "2"]
    };

    factory_from_json_bool_coercion => {
        r#"class Flags {
  bool active;
  Flags._(this.active);
  factory Flags.fromJson(Map<String, dynamic> json) {
    return Flags._(json['active'] == true);
  }
}
void main() {
  print(Flags.fromJson({'active': true}).active);
}"#,
        ["true"]
    };

    factory_from_json_optional_string_null => {
        r#"class Note {
  String? body;
  Note._(this.body);
  factory Note.fromJson(Map<String, dynamic> json) {
    return Note._(json['body']);
  }
}
void main() {
  print(Note.fromJson({}).body);
}"#,
        ["null"]
    };

    factory_implementation_class_hides_private_ctor => {
        r#"class Hidden {
  int code;
  Hidden._(this.code);
  factory Hidden.open(int c) {
    return Hidden._(c);
  }
}
void main() {
  print(Hidden.open(42).code);
}"#,
        ["42"]
    };

    factory_named_from_list_first_element => {
        r#"class Head {
  int value;
  Head(this.value);
  factory Head.fromList(List<int> items) {
    return Head(items.first);
  }
}
void main() {
  print(Head.fromList([9, 8, 7]).value);
}"#,
        ["9"]
    };

    factory_redirect_with_default_named_param => {
        r#"class Id {
  final int value;
  Id._(this.value);
  factory Id({int seed = 1}) => Id._(seed);
}
void main() {
  print(Id().value);
}"#,
        ["1"]
    };

    factory_redirect_with_explicit_named_param => {
        r#"class Id {
  final int value;
  Id._(this.value);
  factory Id({int seed = 1}) => Id._(seed);
}
void main() {
  print(Id(seed: 99).value);
}"#,
        ["99"]
    };

    factory_validates_range_before_construct => {
        r#"class Percent {
  int value;
  Percent._(this.value);
  factory Percent(int v) {
    assert(v >= 0 && v <= 100);
    return Percent._(v);
  }
}
void main() {
  print(Percent(75).value);
}"#,
        ["75"]
    };

    factory_from_json_string_to_int_parse => {
        r#"class Score {
  int points;
  Score._(this.points);
  factory Score.fromJson(Map<String, dynamic> json) {
    return Score._(int.parse(json['points']));
  }
}
void main() {
  print(Score.fromJson({'points': '42'}).points);
}"#,
        ["42"]
    };

    factory_creates_empty_via_private => {
        r#"class Bag {
  List<int> items;
  Bag._(this.items);
  factory Bag.empty() {
    return Bag._([]);
  }
}
void main() {
  print(Bag.empty().items.length);
}"#,
        ["0"]
    };

    factory_singleton_lazy_initialization => {
        r#"class Db {
  static Db? _conn;
  bool ready = true;
  Db._();
  factory Db.connect() {
    _conn ??= Db._();
    return _conn!;
  }
}
void main() {
  print(Db.connect().ready);
}"#,
        ["true"]
    };

    factory_from_json_with_fallback_name => {
        r#"class Label {
  String text;
  Label._(this.text);
  factory Label.fromJson(Map<String, dynamic> json) {
    var t = json['text'] ?? json['label'] ?? 'unknown';
    return Label._(t);
  }
}
void main() {
  print(Label.fromJson({'label': 'ok'}).text);
}"#,
        ["ok"]
    };

    factory_multiple_named_entry_points => {
        r#"class Color {
  int r;
  int g;
  int b;
  Color(this.r, this.g, this.b);
  factory Color.black() {
    return Color(0, 0, 0);
  }
  factory Color.white() {
    return Color(255, 255, 255);
  }
}
void main() {
  print(Color.black().r + Color.white().r);
}"#,
        ["255"]
    };

    factory_returns_subclass_for_special_case => {
        r#"class Animal {
  String kind;
  Animal(this.kind);
  factory Animal.dog() {
    return Dog('dog');
  }
}
class Dog extends Animal {
  Dog(String k) : super(k);
}
void main() {
  print(Animal.dog().kind);
}"#,
        ["dog"]
    };

    factory_with_local_variable_before_return => {
        r#"class Wrap {
  int doubled;
  Wrap._(this.doubled);
  factory Wrap(int n) {
    var d = n * 2;
    return Wrap._(d);
  }
}
void main() {
  print(Wrap(6).doubled);
}"#,
        ["12"]
    };

    factory_from_json_sum_numeric_fields => {
        r#"class Totals {
  int sum;
  Totals._(this.sum);
  factory Totals.fromJson(Map<String, dynamic> json) {
    return Totals._(json['a'] + json['b']);
  }
}
void main() {
  print(Totals.fromJson({'a': 10, 'b': 5}).sum);
}"#,
        ["15"]
    };

    factory_cache_cleared_manually_still_works => {
        r#"class Session {
  static Session? _active;
  int id;
  Session._(this.id);
  factory Session(int id) {
    _active = Session._(id);
    return _active!;
  }
}
void main() {
  print(Session(3).id);
}"#,
        ["3"]
    };

    factory_redirect_chain_to_primary => {
        r#"class Node {
  int depth;
  Node(this.depth);
  Node.root() : depth = 0;
  factory Node.zero() = Node.root;
}
void main() {
  print(Node.zero().depth);
}"#,
        ["0"]
    };

    factory_from_json_extracts_list_first_string => {
        r#"class TagList {
  String primary;
  TagList._(this.primary);
  factory TagList.fromJson(Map<String, dynamic> json) {
    var tags = json['tags'] as List;
    return TagList._(tags.first);
  }
}
void main() {
  print(TagList.fromJson({'tags': ['alpha', 'beta']}).primary);
}"#,
        ["alpha"]
    };

    factory_validates_non_empty_string => {
        r#"class Name {
  String value;
  Name._(this.value);
  factory Name(String v) {
    assert(v.isNotEmpty);
    return Name._(v);
  }
}
void main() {
  print(Name('Zed').value);
}"#,
        ["Zed"]
    };

    factory_bool_gate_selects_implementation => {
        r#"class Result {
  int code;
  Result(this.code);
  factory Result.ok() {
    return Result(0);
  }
  factory Result.err() {
    return Result(1);
  }
}
void main() {
  print(Result.ok().code + Result.err().code);
}"#,
        ["1"]
    };

    factory_from_json_nested_int_access => {
        r#"class Meta {
  int level;
  Meta._(this.level);
  factory Meta.fromJson(Map<String, dynamic> json) {
    return Meta._(json['meta']['level']);
  }
}
void main() {
  print(Meta.fromJson({'meta': {'level': 4}}).level);
}"#,
        ["4"]
    };

    factory_with_assert_on_factory_param_length => {
        r#"class Code {
  String text;
  Code._(this.text);
  factory Code(String t) {
    assert(t.length >= 2);
    return Code._(t);
  }
}
void main() {
  print(Code('xy').text.length);
}"#,
        ["2"]
    };

    factory_parses_enum_like_string => {
        r#"class Mode {
  String value;
  Mode._(this.value);
  factory Mode.parse(String s) {
    if (s == 'on') {
      return Mode._('on');
    }
    return Mode._('off');
  }
}
void main() {
  print(Mode.parse('on').value);
}"#,
        ["on"]
    };

    factory_registry_returns_existing_for_same_id => {
        r#"class Entity {
  static Map<int, Entity> store = {};
  int id;
  Entity._(this.id);
  factory Entity.get(int id) {
    if (store.containsKey(id)) {
      return store[id]!;
    }
    var e = Entity._(id);
    store[id] = e;
    return e;
  }
}
void main() {
  print(Entity.get(1) == Entity.get(1));
}"#,
        ["true"]
    };

    factory_from_json_with_type_cast_list => {
        r#"class Pack {
  List<String> tags;
  Pack._(this.tags);
  factory Pack.fromJson(Map<String, dynamic> json) {
    var raw = json['tags'] as List;
    return Pack._(raw.map((e) => e as String).toList());
  }
}
void main() {
  var p = Pack.fromJson({'tags': ['a', 'b']});
  print(p.tags.join('-'));
}"#,
        ["a-b"]
    };

    factory_combines_two_json_fields => {
        r#"class FullName {
  String full;
  FullName._(this.full);
  factory FullName.fromJson(Map<String, dynamic> json) {
    return FullName._('${json['first']} ${json['last']}');
  }
}
void main() {
  print(FullName.fromJson({'first': 'Ada', 'last': 'Lovelace'}).full);
}"#,
        ["Ada Lovelace"]
    };

    factory_with_early_return_cached_path => {
        r#"class Blob {
  static Blob? _empty;
  int size;
  Blob._(this.size);
  factory Blob.empty() {
    if (_empty != null) {
      return _empty!;
    }
    _empty = Blob._(0);
    return _empty!;
  }
}
void main() {
  print(Blob.empty().size);
}"#,
        ["0"]
    };

    factory_redirect_unnamed_to_private_named => {
        r#"class Key {
  String value;
  Key._(this.value);
  factory Key(String v) => Key._(v);
}
void main() {
  print(Key('id').value);
}"#,
        ["id"]
    };
}
