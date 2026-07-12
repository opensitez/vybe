//! Static fields, static methods, and static getters/setters.

dart_cases! {
    static_int_field_read_by_class_name => {
        r#"class Config {
  static int port = 8080;
}
void main() {
  print(Config.port);
}"#,
        ["8080"]
    };

    static_string_field_read => {
        r#"class Labels {
  static String app = 'vybe';
}
void main() {
  print(Labels.app);
}"#,
        ["vybe"]
    };

    static_bool_field_read => {
        r#"class Flags {
  static bool debug = true;
}
void main() {
  print(Flags.debug);
}"#,
        ["true"]
    };

    static_field_mutation_increments => {
        r#"class Hits {
  static int count = 0;
}
void main() {
  Hits.count = Hits.count + 1;
  Hits.count = Hits.count + 1;
  print(Hits.count);
}"#,
        ["2"]
    };

    static_method_returns_sum => {
        r#"class Add {
  static int sum(int a, int b) {
    return a + b;
  }
}
void main() {
  print(Add.sum(4, 9));
}"#,
        ["13"]
    };

    static_method_with_single_arg => {
        r#"class Square {
  static int sq(int n) {
    return n * n;
  }
}
void main() {
  print(Square.sq(6));
}"#,
        ["36"]
    };

    static_method_arrow_body => {
        r#"class Id {
  static int next() => 99;
}
void main() {
  print(Id.next());
}"#,
        ["99"]
    };

    static_getter_exposes_private_field => {
        r#"class Build {
  static int _rev = 7;
  static int get rev => _rev;
}
void main() {
  print(Build.rev);
}"#,
        ["7"]
    };

    static_setter_updates_private_field => {
        r#"class Limit {
  static int _max = 1;
  static int get max => _max;
  static set max(int v) {
    _max = v;
  }
}
void main() {
  Limit.max = 10;
  print(Limit.max);
}"#,
        ["10"]
    };

    static_method_calls_other_static_method => {
        r#"class Chain {
  static int a(int n) {
    return n + 1;
  }
  static int b(int n) {
    return a(n) * 2;
  }
}
void main() {
  print(Chain.b(3));
}"#,
        ["8"]
    };

    static_const_int_field => {
        r#"class Math {
  static const int piApprox = 3;
}
void main() {
  print(Math.piApprox);
}"#,
        ["3"]
    };

    static_const_string_field => {
        r#"class Meta {
  static const String version = '1.0';
}
void main() {
  print(Meta.version);
}"#,
        ["1.0"]
    };

    static_method_returns_string => {
        r#"class Greet {
  static String hello(String name) {
    return 'hi $name';
  }
}
void main() {
  print(Greet.hello('Ann'));
}"#,
        ["hi Ann"]
    };

    static_field_read_from_instance_method => {
        r#"class Reader {
  static int shared = 11;
  int read() {
    return shared;
  }
}
void main() {
  print(Reader().read());
}"#,
        ["11"]
    };

    static_method_mutates_static_field => {
        r#"class Counter {
  static int n = 0;
  static void inc() {
    n = n + 1;
  }
}
void main() {
  Counter.inc();
  Counter.inc();
  print(Counter.n);
}"#,
        ["2"]
    };

    static_getter_computes_from_static_field => {
        r#"class Scale {
  static int base = 5;
  static int get doubled => base * 2;
}
void main() {
  print(Scale.doubled);
}"#,
        ["10"]
    };

    static_method_with_default_optional_param => {
        r#"class Opt {
  static int mul(int a, [int b = 2]) {
    return a * b;
  }
}
void main() {
  print(Opt.mul(7));
}"#,
        ["14"]
    };

    static_method_with_named_optional_param => {
        r#"class Named {
  static int add({int a = 0, int b = 0}) {
    return a + b;
  }
}
void main() {
  print(Named.add(a: 3, b: 4));
}"#,
        ["7"]
    };

    static_field_zero_initial => {
        r#"class Zero {
  static int value = 0;
}
void main() {
  print(Zero.value);
}"#,
        ["0"]
    };

    static_method_returns_bool => {
        r#"class Check {
  static bool isPositive(int n) {
    return n > 0;
  }
}
void main() {
  print(Check.isPositive(3));
}"#,
        ["true"]
    };

    static_only_class_no_instance_needed => {
        r#"class Util {
  static int id = 42;
  static int getId() {
    return id;
  }
}
void main() {
  print(Util.getId());
}"#,
        ["42"]
    };

    static_method_three_arg_product => {
        r#"class Prod {
  static int triple(int a, int b, int c) {
    return a * b * c;
  }
}
void main() {
  print(Prod.triple(2, 3, 4));
}"#,
        ["24"]
    };

    static_field_assignment_from_method => {
        r#"class Store {
  static int slot = 1;
  static void save(int v) {
    slot = v;
  }
}
void main() {
  Store.save(88);
  print(Store.slot);
}"#,
        ["88"]
    };

    static_getter_string_from_static_int => {
        r#"class Code {
  static int num = 404;
  static String get label => 'c$num';
}
void main() {
  print(Code.label);
}"#,
        ["c404"]
    };

    static_method_max_of_two => {
        r#"class Max {
  static int bigger(int a, int b) {
    return a > b ? a : b;
  }
}
void main() {
  print(Max.bigger(12, 8));
}"#,
        ["12"]
    };

    static_field_negative_value => {
        r#"class Temp {
  static int offset = -5;
}
void main() {
  print(Temp.offset);
}"#,
        ["-5"]
    };

    static_method_invoked_from_instance_context => {
        r#"class Mixed {
  static int stat() {
    return 3;
  }
  int useStat() {
    return stat();
  }
}
void main() {
  print(Mixed().useStat());
}"#,
        ["3"]
    };

    static_compound_plus_assign_field => {
        r#"class Acc {
  static int total = 10;
}
void main() {
  Acc.total += 5;
  print(Acc.total);
}"#,
        ["15"]
    };

    static_method_returns_negated_arg => {
        r#"class Neg {
  static int flip(int n) {
    return -n;
  }
}
void main() {
  print(Neg.flip(9));
}"#,
        ["-9"]
    };

    static_getter_and_instance_method_coexist => {
        r#"class Dual {
  static int get base => 2;
  int inst() {
    return 1;
  }
}
void main() {
  print(Dual().inst() + Dual.base);
}"#,
        ["3"]
    };

    static_final_like_const_int => {
        r#"class Limits {
  static const int maxItems = 100;
}
void main() {
  print(Limits.maxItems);
}"#,
        ["100"]
    };

    static_method_string_length => {
        r#"class Text {
  static int len(String s) {
    return s.length;
  }
}
void main() {
  print(Text.len('dart'));
}"#,
        ["4"]
    };

    static_field_updated_twice_in_main => {
        r#"class State {
  static int mode = 0;
}
void main() {
  State.mode = 1;
  State.mode = 2;
  print(State.mode);
}"#,
        ["2"]
    };

    static_method_identity => {
        r#"class Echo {
  static int pass(int n) {
    return n;
  }
}
void main() {
  print(Echo.pass(55));
}"#,
        ["55"]
    };

    static_getter_bool_flag => {
        r#"class Switch {
  static bool _on = false;
  static bool get on => _on;
}
void main() {
  print(Switch.on);
}"#,
        ["false"]
    };
}
