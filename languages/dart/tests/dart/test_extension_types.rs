//! Dart 3 extension types: int/String representation, representation field
//! access, instance methods, and static helpers on extension types.

dart_cases! {
    extension_type_int_wraps_value => {
        r#"extension type UserId(int id) {}
void main() {
  UserId uid = UserId(42);
  print(uid.id);
}"#,
        ["42"]
    };

    extension_type_string_wraps_value => {
        r#"extension type Tag(String label) {}
void main() {
  Tag t = Tag('dart');
  print(t.label);
}"#,
        ["dart"]
    };

    extension_type_int_method_doubles_representation => {
        r#"extension type Meters(int value) {
  int toCentimeters() {
    return value * 100;
  }
}
void main() {
  Meters m = Meters(3);
  print(m.toCentimeters());
}"#,
        ["300"]
    };

    extension_type_string_method_uppercases => {
        r#"extension type Name(String value) {
  String shout() {
    return value.toUpperCase();
  }
}
void main() {
  Name n = Name('hello');
  print(n.shout());
}"#,
        ["HELLO"]
    };

    extension_type_int_getter_is_positive => {
        r#"extension type Score(int points) {
  bool get isPositive {
    return points > 0;
  }
}
void main() {
  Score s = Score(5);
  print(s.isPositive);
}"#,
        ["true"]
    };

    extension_type_int_getter_is_zero => {
        r#"extension type Score(int points) {
  bool get isZero {
    return points == 0;
  }
}
void main() {
  Score s = Score(0);
  print(s.isZero);
}"#,
        ["true"]
    };

    extension_type_string_getter_length => {
        r#"extension type Label(String text) {
  int get length {
    return text.length;
  }
}
void main() {
  Label l = Label('dart');
  print(l.length);
}"#,
        ["4"]
    };

    extension_type_string_getter_is_empty => {
        r#"extension type Label(String text) {
  bool get isEmpty {
    return text.isEmpty;
  }
}
void main() {
  Label l = Label('');
  print(l.isEmpty);
}"#,
        ["true"]
    };

    extension_type_int_increment_method => {
        r#"extension type Counter(int count) {
  Counter increment() {
    return Counter(count + 1);
  }
}
void main() {
  Counter c = Counter(9);
  print(c.increment().count);
}"#,
        ["10"]
    };

    extension_type_int_decrement_method => {
        r#"extension type Counter(int count) {
  Counter decrement() {
    return Counter(count - 1);
  }
}
void main() {
  Counter c = Counter(5);
  print(c.decrement().count);
}"#,
        ["4"]
    };

    extension_type_string_append_suffix => {
        r#"extension type Slug(String value) {
  Slug withSuffix(String suffix) {
    return Slug(value + suffix);
  }
}
void main() {
  Slug s = Slug('page');
  print(s.withSuffix('-1').value);
}"#,
        ["page-1"]
    };

    extension_type_string_prepend_prefix => {
        r#"extension type Slug(String value) {
  Slug withPrefix(String prefix) {
    return Slug(prefix + value);
  }
}
void main() {
  Slug s = Slug('home');
  print(s.withPrefix('/').value);
}"#,
        ["/home"]
    };

    extension_type_int_add_two_meters => {
        r#"extension type Meters(int value) {
  Meters add(Meters other) {
    return Meters(value + other.value);
  }
}
void main() {
  Meters a = Meters(2);
  Meters b = Meters(3);
  print(a.add(b).value);
}"#,
        ["5"]
    };

    extension_type_int_multiply_by_scalar => {
        r#"extension type Meters(int value) {
  Meters scale(int factor) {
    return Meters(value * factor);
  }
}
void main() {
  Meters m = Meters(4);
  print(m.scale(3).value);
}"#,
        ["12"]
    };

    extension_type_int_compare_greater => {
        r#"extension type Age(int years) {
  bool isOlderThan(Age other) {
    return years > other.years;
  }
}
void main() {
  Age a = Age(30);
  Age b = Age(20);
  print(a.isOlderThan(b));
}"#,
        ["true"]
    };

    extension_type_int_compare_not_greater => {
        r#"extension type Age(int years) {
  bool isOlderThan(Age other) {
    return years > other.years;
  }
}
void main() {
  Age a = Age(10);
  Age b = Age(20);
  print(a.isOlderThan(b));
}"#,
        ["false"]
    };

    extension_type_string_contains_substring => {
        r#"extension type Path(String value) {
  bool containsPart(String part) {
    return value.contains(part);
  }
}
void main() {
  Path p = Path('/usr/bin');
  print(p.containsPart('bin'));
}"#,
        ["true"]
    };

    extension_type_string_starts_with_prefix => {
        r#"extension type Path(String value) {
  bool starts(String prefix) {
    return value.startsWith(prefix);
  }
}
void main() {
  Path p = Path('hello.dart');
  print(p.starts('hello'));
}"#,
        ["true"]
    };

    extension_type_int_abs_on_negative => {
        r#"extension type Offset(int delta) {
  int absValue() {
    return delta.abs();
  }
}
void main() {
  Offset o = Offset(-7);
  print(o.absValue());
}"#,
        ["7"]
    };

    extension_type_int_abs_on_positive => {
        r#"extension type Offset(int delta) {
  int absValue() {
    return delta.abs();
  }
}
void main() {
  Offset o = Offset(7);
  print(o.absValue());
}"#,
        ["7"]
    };

    extension_type_string_trim_whitespace => {
        r#"extension type Token(String raw) {
  String trimmed() {
    return raw.trim();
  }
}
void main() {
  Token t = Token('  ok  ');
  print(t.trimmed());
}"#,
        ["ok"]
    };

    extension_type_string_split_to_list => {
        r#"extension type CsvLine(String line) {
  List<String> cells() {
    return line.split(',');
  }
}
void main() {
  CsvLine row = CsvLine('a,b,c');
  print(row.cells().length);
  print(row.cells()[1]);
}"#,
        ["3", "b"]
    };

    extension_type_int_is_even => {
        r#"extension type Number(int n) {
  bool get isEven {
    return n % 2 == 0;
  }
}
void main() {
  Number num = Number(8);
  print(num.isEven);
}"#,
        ["true"]
    };

    extension_type_int_is_odd => {
        r#"extension type Number(int n) {
  bool get isOdd {
    return n % 2 == 1;
  }
}
void main() {
  Number num = Number(7);
  print(num.isOdd);
}"#,
        ["true"]
    };

    extension_type_int_to_string_via_method => {
        r#"extension type Port(int number) {
  String label() {
    return 'port-$number';
  }
}
void main() {
  Port p = Port(8080);
  print(p.label());
}"#,
        ["port-8080"]
    };

    extension_type_string_parse_to_int => {
        r#"extension type DigitString(String digits) {
  int asInt() {
    return int.parse(digits);
  }
}
void main() {
  DigitString d = DigitString('456');
  print(d.asInt());
}"#,
        ["456"]
    };

    extension_type_int_static_zero_factory => {
        r#"extension type Count(int value) {
  static Count zero() {
    return Count(0);
  }
}
void main() {
  Count c = Count.zero();
  print(c.value);
}"#,
        ["0"]
    };

    extension_type_int_static_from_string => {
        r#"extension type Count(int value) {
  static Count parse(String s) {
    return Count(int.parse(s));
  }
}
void main() {
  Count c = Count.parse('12');
  print(c.value);
}"#,
        ["12"]
    };

    extension_type_string_static_empty => {
        r#"extension type Note(String text) {
  static Note empty() {
    return Note('');
  }
}
void main() {
  Note n = Note.empty();
  print(n.text);
  print(n.text.isEmpty);
}"#,
        ["", "true"]
    };

    extension_type_string_static_concat => {
        r#"extension type Note(String text) {
  static Note join(Note a, Note b) {
    return Note(a.text + b.text);
  }
}
void main() {
  Note n = Note.join(Note('a'), Note('b'));
  print(n.text);
}"#,
        ["ab"]
    };

    extension_type_int_equality_same_value => {
        r#"extension type Id(int value) {
  bool sameAs(Id other) {
    return value == other.value;
  }
}
void main() {
  Id a = Id(5);
  Id b = Id(5);
  print(a.sameAs(b));
}"#,
        ["true"]
    };

    extension_type_int_equality_different_value => {
        r#"extension type Id(int value) {
  bool sameAs(Id other) {
    return value == other.value;
  }
}
void main() {
  Id a = Id(5);
  Id b = Id(6);
  print(a.sameAs(b));
}"#,
        ["false"]
    };

    extension_type_string_equality_case_sensitive => {
        r#"extension type Code(String value) {
  bool matches(Code other) {
    return value == other.value;
  }
}
void main() {
  Code a = Code('ABC');
  Code b = Code('ABC');
  print(a.matches(b));
}"#,
        ["true"]
    };

    extension_type_int_negate_representation => {
        r#"extension type Temperature(int celsius) {
  Temperature negate() {
    return Temperature(-celsius);
  }
}
void main() {
  Temperature t = Temperature(20);
  print(t.negate().celsius);
}"#,
        ["-20"]
    };

    extension_type_int_clamp_between_bounds => {
        r#"extension type Percent(int value) {
  Percent clamped() {
    if (value < 0) return Percent(0);
    if (value > 100) return Percent(100);
    return Percent(value);
  }
}
void main() {
  Percent p = Percent(150);
  print(p.clamped().value);
}"#,
        ["100"]
    };

    extension_type_int_clamp_low_bound => {
        r#"extension type Percent(int value) {
  Percent clamped() {
    if (value < 0) return Percent(0);
    if (value > 100) return Percent(100);
    return Percent(value);
  }
}
void main() {
  Percent p = Percent(-5);
  print(p.clamped().value);
}"#,
        ["0"]
    };

    extension_type_string_char_at_index => {
        r#"extension type Word(String value) {
  String charAt(int index) {
    return value[index];
  }
}
void main() {
  Word w = Word('dart');
  print(w.charAt(0));
  print(w.charAt(3));
}"#,
        ["d", "t"]
    };

    extension_type_string_replace_substring => {
        r#"extension type Sentence(String value) {
  Sentence replaceAll(String from, String to) {
    return Sentence(value.replaceAll(from, to));
  }
}
void main() {
  Sentence s = Sentence('foo bar foo');
  print(s.replaceAll('foo', 'baz').value);
}"#,
        ["baz bar baz"]
    };

    extension_type_int_bitwise_and => {
        r#"extension type Flags(int bits) {
  Flags and(Flags other) {
    return Flags(bits & other.bits);
  }
}
void main() {
  Flags a = Flags(12);
  Flags b = Flags(10);
  print(a.and(b).bits);
}"#,
        ["8"]
    };

    extension_type_int_bitwise_or => {
        r#"extension type Flags(int bits) {
  Flags or(Flags other) {
    return Flags(bits | other.bits);
  }
}
void main() {
  Flags a = Flags(1);
  Flags b = Flags(2);
  print(a.or(b).bits);
}"#,
        ["3"]
    };

    extension_type_int_shift_left => {
        r#"extension type Bits(int value) {
  Bits shiftLeft(int n) {
    return Bits(value << n);
  }
}
void main() {
  Bits b = Bits(3);
  print(b.shiftLeft(2).value);
}"#,
        ["12"]
    };

    extension_type_string_substring_range => {
        r#"extension type Text(String value) {
  String slice(int start, int end) {
    return value.substring(start, end);
  }
}
void main() {
  Text t = Text('hello');
  print(t.slice(1, 4));
}"#,
        ["ell"]
    };

    extension_type_int_divide_representation => {
        r#"extension type Quantity(int amount) {
  int divideBy(int divisor) {
    return amount ~/ divisor;
  }
}
void main() {
  Quantity q = Quantity(17);
  print(q.divideBy(5));
}"#,
        ["3"]
    };

    extension_type_int_modulo_representation => {
        r#"extension type Quantity(int amount) {
  int remainder(int divisor) {
    return amount % divisor;
  }
}
void main() {
  Quantity q = Quantity(17);
  print(q.remainder(5));
}"#,
        ["2"]
    };

    extension_type_string_repeat_n_times => {
        r#"extension type Pattern(String value) {
  String repeat(int times) {
    return value * times;
  }
}
void main() {
  Pattern p = Pattern('ab');
  print(p.repeat(3));
}"#,
        ["ababab"]
    };

    extension_type_int_max_of_two => {
        r#"extension type Measure(int value) {
  Measure max(Measure other) {
    return Measure(value > other.value ? value : other.value);
  }
}
void main() {
  Measure a = Measure(4);
  Measure b = Measure(9);
  print(a.max(b).value);
}"#,
        ["9"]
    };

    extension_type_int_chain_two_methods => {
        r#"extension type Step(int n) {
  Step next() {
    return Step(n + 1);
  }
  Step doubleStep() {
    return Step(n + 2);
  }
}
void main() {
  Step s = Step(1);
  print(s.next().doubleStep().n);
}"#,
        ["4"]
    };

    extension_type_string_chain_two_methods => {
        r#"extension type Raw(String value) {
  Raw trimSpaces() {
    return Raw(value.trim());
  }
  Raw addBang() {
    return Raw(value + '!');
  }
}
void main() {
  Raw r = Raw('  hi  ');
  print(r.trimSpaces().addBang().value);
}"#,
        ["hi!"]
    };
}
