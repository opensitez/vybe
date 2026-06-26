//! Typedef aliases for function types and generic type aliases.

dart_cases! {
    typedef_int_unary_function_invoked => {
        r#"typedef IntFn = int Function(int);
int triple(int n) {
  return n * 3;
}
void main() {
  IntFn fn = triple;
  print(fn(4));
}"#,
        ["12"]
    };

    typedef_int_binary_function_sum => {
        r#"typedef IntOp = int Function(int, int);
int add(int a, int b) {
  return a + b;
}
void main() {
  IntOp op = add;
  print(op(7, 5));
}"#,
        ["12"]
    };

    typedef_void_callback_invoked => {
        r#"typedef Logger = void Function(String);
var buffer = <String>[];
void capture(String msg) {
  buffer.add(msg);
}
void main() {
  Logger log = capture;
  log('ready');
  print(buffer.join(','));
}"#,
        ["ready"]
    };

    typedef_bool_predicate_returns_true => {
        r#"typedef Filter = bool Function(int);
bool isPositive(int n) {
  return n > 0;
}
void main() {
  Filter pred = isPositive;
  print(pred(3));
}"#,
        ["true"]
    };

    typedef_bool_predicate_returns_false => {
        r#"typedef Filter = bool Function(int);
bool isPositive(int n) {
  return n > 0;
}
void main() {
  Filter pred = isPositive;
  print(pred(-1));
}"#,
        ["false"]
    };

    typedef_double_binary_multiplies => {
        r#"typedef Scale = double Function(double, double);
double mul(double a, double b) {
  return a * b;
}
void main() {
  Scale fn = mul;
  print(fn(2.5, 4.0));
}"#,
        ["10.0"]
    };

    typedef_nullary_factory_returns_string => {
        r#"typedef Factory = String Function();
String makeLabel() {
  return 'vybe';
}
void main() {
  Factory f = makeLabel;
  print(f());
}"#,
        ["vybe"]
    };

    typedef_passed_as_function_argument => {
        r#"typedef Mapper = int Function(int);
int applyTwice(int value, Mapper fn) {
  return fn(fn(value));
}
int inc(int n) {
  return n + 1;
}
void main() {
  print(applyTwice(3, inc));
}"#,
        ["5"]
    };

    typedef_returned_from_helper => {
        r#"typedef Op = int Function(int);
Op makeAdder(int delta) {
  return (int n) => n + delta;
}
void main() {
  Op add5 = makeAdder(5);
  print(add5(10));
}"#,
        ["15"]
    };

    typedef_assigned_from_closure => {
        r#"typedef Doubler = int Function(int);
void main() {
  Doubler fn = (int n) => n * 2;
  print(fn(6));
}"#,
        ["12"]
    };

    typedef_stored_in_list_and_invoked => {
        r#"typedef Step = int Function(int);
int addOne(int n) {
  return n + 1;
}
int addTwo(int n) {
  return n + 2;
}
void main() {
  List<Step> steps = [addOne, addTwo];
  print(steps[0](5));
  print(steps[1](5));
}"#,
        ["6", "7"]
    };

    typedef_class_field_invoked => {
        r#"typedef Handler = String Function(String);
class Greeter {
  Handler handler;
  Greeter(this.handler);
  String run(String name) {
    return handler(name);
  }
}
void main() {
  var g = Greeter((name) => 'hi $name');
  print(g.run('dart'));
}"#,
        ["hi dart"]
    };

    typedef_string_transformer_uppercases => {
        r#"typedef Transformer = String Function(String);
String shout(String s) {
  return s.toUpperCase();
}
void main() {
  Transformer t = shout;
  print(t('go'));
}"#,
        ["GO"]
    };

    typedef_comparator_orders_descending => {
        r#"typedef Compare = int Function(int, int);
int desc(int a, int b) {
  return b.compareTo(a);
}
void main() {
  Compare cmp = desc;
  print(cmp(2, 9));
}"#,
        ["1"]
    };

    typedef_chained_through_two_wrappers => {
        r#"typedef Fn = int Function(int);
int wrap(int n, Fn fn) {
  return fn(n) + 1;
}
int square(int n) {
  return n * n;
}
void main() {
  print(wrap(3, square));
}"#,
        ["10"]
    };

    typedef_three_param_sum => {
        r#"typedef Sum3 = int Function(int, int, int);
int sum3(int a, int b, int c) {
  return a + b + c;
}
void main() {
  Sum3 fn = sum3;
  print(fn(1, 2, 3));
}"#,
        ["6"]
    };

    typedef_optional_wrapper_with_default => {
        r#"typedef MaybeFn = int Function(int);
int runOrZero(MaybeFn? fn, int input) {
  if (fn == null) {
    return 0;
  }
  return fn(input);
}
int doubleIt(int n) {
  return n * 2;
}
void main() {
  print(runOrZero(doubleIt, 4));
  print(runOrZero(null, 4));
}"#,
        ["8", "0"]
    };

    typedef_map_value_callback => {
        r#"typedef Callback = int Function(int);
int runAll(Map<String, Callback> table, int seed) {
  var total = 0;
  table.forEach((key, fn) {
    total += fn(seed);
  });
  return total;
}
void main() {
  var table = <String, Callback>{
    'a': (n) => n + 1,
    'b': (n) => n * 2,
  };
  print(runAll(table, 3));
}"#,
        ["10"]
    };

    typedef_dynamic_function_accepts_any => {
        r#"typedef DynFn = dynamic Function(dynamic);
dynamic echo(dynamic value) {
  return value;
}
void main() {
  DynFn fn = echo;
  print(fn('x'));
  print(fn(9));
}"#,
        ["x", "9"]
    };

    typedef_named_target_function => {
        r#"typedef Formatter = String Function({required String prefix, required String body});
String format({required String prefix, required String body}) {
  return '$prefix:$body';
}
void main() {
  Formatter fn = format;
  print(fn(prefix: 'id', body: '42'));
}"#,
        ["id:42"]
    };

    generic_typedef_list_alias_length => {
        r#"typedef IntList = List<int>;
void main() {
  IntList nums = [10, 20, 30];
  print(nums.length);
}"#,
        ["3"]
    };

    generic_typedef_string_list_join => {
        r#"typedef StringList = List<String>;
void main() {
  StringList words = ['a', 'b', 'c'];
  print(words.join('-'));
}"#,
        ["a-b-c"]
    };

    generic_typedef_map_alias_keys => {
        r#"typedef StringIntMap = Map<String, int>;
void main() {
  StringIntMap scores = {'a': 1, 'b': 2};
  print(scores.keys.length);
  print(scores['b']);
}"#,
        ["2", "2"]
    };

    generic_typedef_set_alias_length => {
        r#"typedef IntSet = Set<int>;
void main() {
  IntSet values = {1, 2, 2, 3};
  print(values.length);
}"#,
        ["3"]
    };

    generic_typedef_callback_with_type_param => {
        r#"typedef Callback<T> = void Function(T);
var seen = <String>[];
void remember(String value) {
  seen.add(value);
}
void main() {
  Callback<String> cb = remember;
  cb('dart');
  print(seen.join(','));
}"#,
        ["dart"]
    };

    generic_typedef_mapper_transforms_list => {
        r#"typedef Mapper<T, R> = R Function(T);
List<R> mapAll<T, R>(List<T> items, Mapper<T, R> fn) {
  var out = <R>[];
  for (var item in items) {
    out.add(fn(item));
  }
  return out;
}
void main() {
  var lengths = mapAll(['a', 'bb'], (String s) => s.length);
  print(lengths.join(','));
}"#,
        ["1,2"]
    };

    generic_typedef_predicate_filters_values => {
        r#"typedef Predicate<T> = bool Function(T);
List<T> keepIf<T>(List<T> items, Predicate<T> pred) {
  var out = <T>[];
  for (var item in items) {
    if (pred(item)) {
      out.add(item);
    }
  }
  return out;
}
void main() {
  var evens = keepIf([1, 2, 3, 4], (n) => n % 2 == 0);
  print(evens.join(','));
}"#,
        ["2,4"]
    };

    generic_typedef_pair_function_type => {
        r#"typedef PairFn<A, B> = B Function(A);
String label(int code) {
  return 'code-$code';
}
void main() {
  PairFn<int, String> fn = label;
  print(fn(7));
}"#,
        ["code-7"]
    };

    generic_typedef_factory_creates_list => {
        r#"typedef ListFactory<T> = List<T> Function();
List<int> makeInts() {
  return [1, 2, 3];
}
void main() {
  ListFactory<int> factory = makeInts;
  print(factory().join(','));
}"#,
        ["1,2,3"]
    };

    generic_typedef_iterable_alias_first => {
        r#"typedef NumIterable = Iterable<int>;
void main() {
  NumIterable values = [5, 6, 7];
  print(values.first);
}"#,
        ["5"]
    };

    generic_typedef_nested_mapper_to_length => {
        r#"typedef ToLength = int Function(String);
typedef StringList = List<String>;
int totalChars(StringList words, ToLength measure) {
  var sum = 0;
  for (var word in words) {
    sum += measure(word);
  }
  return sum;
}
void main() {
  print(totalChars(['a', 'bb', 'ccc'], (s) => s.length));
}"#,
        ["6"]
    };

    generic_typedef_reducer_sums_ints => {
        r#"typedef Reducer<T> = T Function(T, T);
int foldLeft(List<int> items, Reducer<int> combine, int seed) {
  var acc = seed;
  for (var item in items) {
    acc = combine(acc, item);
  }
  return acc;
}
void main() {
  print(foldLeft([1, 2, 3], (a, b) => a + b, 0));
}"#,
        ["6"]
    };

    generic_typedef_equality_checker => {
        r#"typedef Eq<T> = bool Function(T, T);
bool allEqual<T>(List<T> items, Eq<T> same) {
  if (items.isEmpty) {
    return true;
  }
  var first = items.first;
  for (var item in items) {
    if (!same(item, first)) {
      return false;
    }
  }
  return true;
}
void main() {
  print(allEqual(['x', 'x', 'x'], (a, b) => a == b));
  print(allEqual(['x', 'y'], (a, b) => a == b));
}"#,
        ["true", "false"]
    };

    generic_typedef_optional_result_mapper => {
        r#"typedef Parser<T> = T? Function(String);
int? parseInt(String raw) {
  if (raw == '42') {
    return 42;
  }
  return null;
}
void main() {
  Parser<int> parse = parseInt;
  print(parse('42'));
  print(parse('nope') == null);
}"#,
        ["42", "true"]
    };

    generic_typedef_callback_list_invocation_order => {
        r#"typedef Consumer<T> = void Function(T);
void main() {
  var log = <int>[];
  List<Consumer<int>> steps = [
    (n) => log.add(1),
    (n) => log.add(2),
    (n) => log.add(3),
  ];
  for (var step in steps) {
    step(0);
  }
  print(log.join(','));
}"#,
        ["1,2,3"]
    };

    generic_typedef_map_value_list_lookup => {
        r#"typedef IntList = List<int>;
void main() {
  Map<String, IntList> grouped = {
    'evens': [2, 4],
    'odds': [1, 3]
  };
  print(grouped['evens']!.first);
  print(grouped['odds']!.last);
}"#,
        ["2", "3"]
    };

    generic_typedef_comparator_sorts_descending => {
        r#"typedef IntList = List<int>;
typedef CompareFn = int Function(int, int);
void sortDesc(IntList items, CompareFn cmp) {
  items.sort(cmp);
}
void main() {
  IntList nums = [1, 3, 2];
  sortDesc(nums, (a, b) => b.compareTo(a));
  print(nums.join(','));
}"#,
        ["3,2,1"]
    };

    generic_typedef_box_list_nested_length => {
        r#"typedef BoxList<T> = List<List<T>>;
void main() {
  BoxList<int> matrix = [
    [1, 2],
    [3],
  ];
  print(matrix.length);
  print(matrix[0].length);
  print(matrix[1].single);
}"#,
        ["2", "2", "3"]
    };

    generic_typedef_string_map_getter => {
        r#"typedef Lookup = Map<String, String>;
String read(Lookup table, String key) {
  return table[key] ?? 'missing';
}
void main() {
  Lookup labels = {'en': 'hello', 'fr': 'bonjour'};
  print(read(labels, 'en'));
  print(read(labels, 'de'));
}"#,
        ["hello", "missing"]
    };

    generic_typedef_chain_mapper_and_predicate => {
        r#"typedef Mapper<T, R> = R Function(T);
typedef Predicate<R> = bool Function(R);
bool anyMatch<T, R>(List<T> items, Mapper<T, R> map, Predicate<R> pred) {
  for (var item in items) {
    if (pred(map(item))) {
      return true;
    }
  }
  return false;
}
void main() {
  print(anyMatch(['aa', 'b', 'ccc'], (s) => s.length, (len) => len > 2));
  print(anyMatch(['a', 'bb'], (s) => s.length, (len) => len > 3));
}"#,
        ["true", "false"]
    };
}
