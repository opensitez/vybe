//! Named parameters: optional with defaults, required named, reordering,
//! mixing with positional parameters, and constructor named arguments.

dart_cases! {
    named_param_uses_default_when_omitted => {
        r#"void greet({String name = 'World'}) {
  print('Hello $name');
}
void main() {
  greet();
}"#,
        ["Hello World"]
    };

    named_param_overrides_default => {
        r#"void greet({String name = 'World'}) {
  print('Hello $name');
}
void main() {
  greet(name: 'Dart');
}"#,
        ["Hello Dart"]
    };

    required_named_param_must_be_provided => {
        r#"void connect({required String host}) {
  print(host);
}
void main() {
  connect(host: 'localhost');
}"#,
        ["localhost"]
    };

    required_named_int_param => {
        r#"void bind({required int port}) {
  print(port);
}
void main() {
  bind(port: 3000);
}"#,
        ["3000"]
    };

    optional_named_int_default => {
        r#"void listen({int port = 80}) {
  print(port);
}
void main() {
  listen();
}"#,
        ["80"]
    };

    optional_named_int_override => {
        r#"void listen({int port = 80}) {
  print(port);
}
void main() {
  listen(port: 443);
}"#,
        ["443"]
    };

    named_params_reordered_at_call_site => {
        r#"void pair({int a = 0, int b = 0}) {
  print('$a,$b');
}
void main() {
  pair(b: 2, a: 1);
}"#,
        ["1,2"]
    };

    three_named_params_any_order => {
        r#"void triple({int x = 0, int y = 0, int z = 0}) {
  print('$x$y$z');
}
void main() {
  triple(z: 3, x: 1, y: 2);
}"#,
        ["123"]
    };

    positional_then_named_mixed => {
        r#"void log(String level, {String msg = 'empty'}) {
  print('$level:$msg');
}
void main() {
  log('INFO', msg: 'started');
}"#,
        ["INFO:started"]
    };

    positional_then_named_default_only => {
        r#"void log(String level, {String msg = 'empty'}) {
  print('$level:$msg');
}
void main() {
  log('WARN');
}"#,
        ["WARN:empty"]
    };

    required_named_with_optional_default_sibling => {
        r#"void request({required String path, int timeout = 30}) {
  print('$path/$timeout');
}
void main() {
  request(path: '/api');
}"#,
        ["/api/30"]
    };

    required_named_with_optional_override => {
        r#"void request({required String path, int timeout = 30}) {
  print('$path/$timeout');
}
void main() {
  request(path: '/api', timeout: 5);
}"#,
        ["/api/5"]
    };

    multiple_optional_named_all_defaults => {
        r#"void config({bool debug = false, int retries = 1}) {
  print('$debug,$retries');
}
void main() {
  config();
}"#,
        ["false,1"]
    };

    multiple_optional_named_partial_override => {
        r#"void config({bool debug = false, int retries = 1}) {
  print('$debug,$retries');
}
void main() {
  config(debug: true);
}"#,
        ["true,1"]
    };

    named_bool_param_default_false => {
        r#"void toggle({bool on = false}) {
  print(on);
}
void main() {
  toggle();
}"#,
        ["false"]
    };

    named_bool_param_set_true => {
        r#"void toggle({bool on = false}) {
  print(on);
}
void main() {
  toggle(on: true);
}"#,
        ["true"]
    };

    named_string_list_default => {
        r#"void tags({List<String> items = const ['a']}) {
  print(items.join(','));
}
void main() {
  tags();
}"#,
        ["a"]
    };

    named_string_list_override => {
        r#"void tags({List<String> items = const ['a']}) {
  print(items.join(','));
}
void main() {
  tags(items: ['x', 'y']);
}"#,
        ["x,y"]
    };

    constructor_named_params_with_defaults => {
        r#"class User {
  final String name;
  final int age;
  User({this.name = 'anon', this.age = 0});
}
void main() {
  var u = User();
  print('${u.name}:${u.age}');
}"#,
        ["anon:0"]
    };

    constructor_required_named_params => {
        r#"class Server {
  final String host;
  final int port;
  Server({required this.host, required this.port});
}
void main() {
  var s = Server(host: '127.0.0.1', port: 9000);
  print('${s.host}:${s.port}');
}"#,
        ["127.0.0.1:9000"]
    };

    constructor_named_reordered => {
        r#"class Pair {
  final int first;
  final int second;
  Pair({required this.first, required this.second});
}
void main() {
  var p = Pair(second: 2, first: 1);
  print('${p.first},${p.second}');
}"#,
        ["1,2"]
    };

    named_params_in_method_on_class => {
        r#"class Printer {
  void show({String text = 'none'}) {
    print(text);
  }
}
void main() {
  Printer().show(text: 'line');
}"#,
        ["line"]
    };

    forwarding_named_to_inner_function => {
        r#"void inner({required String key, int n = 0}) {
  print('$key=$n');
}
void outer({required String key, int n = 0}) {
  inner(key: key, n: n);
}
void main() {
  outer(key: 'id', n: 7);
}"#,
        ["id=7"]
    };

    named_double_param_with_default => {
        r#"void scale({double factor = 1.0}) {
  print(factor * 10);
}
void main() {
  scale();
}"#,
        ["10.0"]
    };

    named_double_param_override => {
        r#"void scale({double factor = 1.0}) {
  print(factor * 10);
}
void main() {
  scale(factor: 2.5);
}"#,
        ["25.0"]
    };

    two_required_named_both_provided => {
        r#"void move({required int dx, required int dy}) {
  print('$dx,$dy');
}
void main() {
  move(dx: 3, dy: -1);
}"#,
        ["3,-1"]
    };

    named_after_positional_two_defaults => {
        r#"void fmt(String prefix, {String suffix = '!'}) {
  print('$prefix$suffix');
}
void main() {
  fmt('hi');
}"#,
        ["hi!"]
    };

    named_after_positional_with_override => {
        r#"void fmt(String prefix, {String suffix = '!'}) {
  print('$prefix$suffix');
}
void main() {
  fmt('hi', suffix: '?');
}"#,
        ["hi?"]
    };

    optional_named_string_empty_default => {
        r#"void label({String title = ''}) {
  print(title.isEmpty ? 'blank' : title);
}
void main() {
  label();
}"#,
        ["blank"]
    };

    optional_named_string_nonempty => {
        r#"void label({String title = ''}) {
  print(title.isEmpty ? 'blank' : title);
}
void main() {
  label(title: 'home');
}"#,
        ["home"]
    };

    named_params_compute_sum => {
        r#"int sum({int a = 0, int b = 0, int c = 0}) {
  return a + b + c;
}
void main() {
  print(sum(b: 5, c: 7));
}"#,
        ["12"]
    };

    named_params_with_nullable_optional => {
        r#"void show({String? label = 'default'}) {
  print(label ?? 'null-label');
}
void main() {
  show();
}"#,
        ["default"]
    };

    named_params_nullable_explicit_null => {
        r#"void show({String? label}) {
  print(label ?? 'null-label');
}
void main() {
  show(label: null);
}"#,
        ["null-label"]
    };

    factory_constructor_named_params => {
        r#"class Id {
  final int value;
  Id._(this.value);
  factory Id({int seed = 1}) => Id._(seed);
}
void main() {
  print(Id(seed: 42).value);
}"#,
        ["42"]
    };

    named_params_in_static_method => {
        r#"class Math {
  static int mul({int a = 1, int b = 1}) => a * b;
}
void main() {
  print(Math.mul(a: 6, b: 7));
}"#,
        ["42"]
    };
}
