//! Generic classes, methods, functions, bounded type parameters, and typed collections.

dart_cases! {
    generic_class_stores_int_value => {
        r#"class Box<T> {
  T value;
  Box(this.value);
}
void main() {
  var b = Box<int>(7);
  print(b.value);
}"#,
        ["7"]
    };

    generic_class_stores_string_value => {
        r#"class Box<T> {
  T value;
  Box(this.value);
}
void main() {
  var b = Box<String>('dart');
  print(b.value);
}"#,
        ["dart"]
    };

    generic_class_type_inference_on_constructor => {
        r#"class Box<T> {
  T value;
  Box(this.value);
}
void main() {
  var b = Box(42);
  print(b.value);
}"#,
        ["42"]
    };

    generic_pair_two_type_parameters => {
        r#"class Pair<A, B> {
  A first;
  B second;
  Pair(this.first, this.second);
}
void main() {
  var p = Pair(1, 'one');
  print(p.first);
  print(p.second);
}"#,
        ["1", "one"]
    };

    generic_pair_swapped_types => {
        r#"class Pair<A, B> {
  A first;
  B second;
  Pair(this.first, this.second);
}
void main() {
  var p = Pair('x', 9);
  print(p.first);
  print(p.second);
}"#,
        ["x", "9"]
    };

    generic_stack_push_and_pop => {
        r#"class Stack<T> {
  List<T> _items = [];
  void push(T item) {
    _items.add(item);
  }
  T pop() {
    return _items.removeLast();
  }
}
void main() {
  var s = Stack<int>();
  s.push(1);
  s.push(2);
  print(s.pop());
}"#,
        ["2"]
    };

    generic_stack_is_empty_getter => {
        r#"class Stack<T> {
  List<T> _items = [];
  bool get isEmpty {
    return _items.isEmpty;
  }
}
void main() {
  var s = Stack<String>();
  print(s.isEmpty);
}"#,
        ["true"]
    };

    generic_method_on_class_maps_items => {
        r#"class Holder<T> {
  List<T> items;
  Holder(this.items);
  List<R> mapAll<R>(R Function(T) fn) {
    return items.map(fn).toList();
  }
}
void main() {
  var h = Holder([1, 2, 3]);
  var doubled = h.mapAll((n) => n * 2);
  print(doubled.join(','));
}"#,
        ["2,4,6"]
    };

    generic_method_on_class_filters_items => {
        r#"class Holder<T> {
  List<T> items;
  Holder(this.items);
  List<T> keepWhere(bool Function(T) test) {
    return items.where(test).toList();
  }
}
void main() {
  var h = Holder([1, 2, 3, 4]);
  var evens = h.keepWhere((n) => n % 2 == 0);
  print(evens.join(','));
}"#,
        ["2,4"]
    };

    top_level_generic_identity_function => {
        r#"T identity<T>(T value) {
  return value;
}
void main() {
  print(identity(99));
}"#,
        ["99"]
    };

    top_level_generic_identity_string => {
        r#"T identity<T>(T value) {
  return value;
}
void main() {
  print(identity('hi'));
}"#,
        ["hi"]
    };

    top_level_generic_swap_pair => {
        r#"Pair<T, U> swap<T, U>(Pair<T, U> p) {
  return Pair(p.second, p.first);
}
class Pair<T, U> {
  T first;
  U second;
  Pair(this.first, this.second);
}
void main() {
  var p = Pair(1, 'a');
  var s = swap(p);
  print(s.first);
  print(s.second);
}"#,
        ["a", "1"]
    };

    generic_function_list_repeat => {
        r#"List<T> repeat<T>(T value, int times) {
  return List.generate(times, (_) => value);
}
void main() {
  var list = repeat(7, 3);
  print(list.join(','));
}"#,
        ["7,7,7"]
    };

    generic_bounded_add_on_int => {
        r#"T addNums<T extends num>(T a, T b) {
  return (a + b) as T;
}
void main() {
  print(addNums(3, 4));
}"#,
        ["7"]
    };

    generic_bounded_add_on_double => {
        r#"T addNums<T extends num>(T a, T b) {
  return (a + b) as T;
}
void main() {
  print(addNums(1.5, 2.5));
}"#,
        ["4.0"]
    };

    generic_comparable_max_int => {
        r#"T maxOf<T extends Comparable<T>>(T a, T b) {
  return a.compareTo(b) >= 0 ? a : b;
}
void main() {
  print(maxOf(3, 9));
}"#,
        ["9"]
    };

    generic_comparable_max_string => {
        r#"T maxOf<T extends Comparable<T>>(T a, T b) {
  return a.compareTo(b) >= 0 ? a : b;
}
void main() {
  print(maxOf('apple', 'banana'));
}"#,
        ["banana"]
    };

    generic_extends_object_container => {
        r#"class Container<T extends Object> {
  T value;
  Container(this.value);
}
void main() {
  var c = Container(42);
  print(c.value);
}"#,
        ["42"]
    };

    typed_list_int_literal => {
        r#"void main() {
  List<int> nums = [1, 2, 3];
  print(nums.length);
  print(nums[1]);
}"#,
        ["3", "2"]
    };

    typed_list_string_join => {
        r#"void main() {
  List<String> words = ['a', 'b', 'c'];
  print(words.join('-'));
}"#,
        ["a-b-c"]
    };

    typed_list_bool_contains => {
        r#"void main() {
  List<bool> flags = [true, false, true];
  print(flags.contains(true));
  print(flags.length);
}"#,
        ["true", "3"]
    };

    typed_empty_list_int => {
        r#"void main() {
  List<int> nums = <int>[];
  print(nums.isEmpty);
  print(nums.length);
}"#,
        ["true", "0"]
    };

    typed_map_string_to_int => {
        r#"void main() {
  Map<String, int> scores = {'Ada': 90, 'Bob': 85};
  print(scores['Ada']);
  print(scores.length);
}"#,
        ["90", "2"]
    };

    typed_map_int_to_string_list => {
        r#"void main() {
  Map<int, List<String>> grouped = {1: ['a'], 2: ['b', 'c']};
  print(grouped[2]!.length);
  print(grouped[2]![0]);
}"#,
        ["2", "b"]
    };

    typed_set_int_deduplicates => {
        r#"void main() {
  Set<int> nums = {1, 2, 2, 3};
  print(nums.length);
  print(nums.contains(2));
}"#,
        ["3", "true"]
    };

    typed_set_string_membership => {
        r#"void main() {
  Set<String> tags = {'dart', 'flutter'};
  print(tags.contains('dart'));
  print(tags.length);
}"#,
        ["true", "2"]
    };

    generic_class_extends_base_with_same_type_param => {
        r#"class Base<T> {
  T value;
  Base(this.value);
}
class Child<T> extends Base<T> {
  Child(T v) : super(v);
}
void main() {
  var c = Child<String>('ok');
  print(c.value);
}"#,
        ["ok"]
    };

    generic_concrete_subclass_fixes_type_arg => {
        r#"class Base<T> {
  T value;
  Base(this.value);
}
class IntBox extends Base<int> {
  IntBox(int v) : super(v);
}
void main() {
  var box = IntBox(15);
  print(box.value);
}"#,
        ["15"]
    };

    generic_nullable_field_on_cache => {
        r#"class Cache<T> {
  T? _value;
  T? get value {
    return _value;
  }
  void store(T v) {
    _value = v;
  }
}
void main() {
  var c = Cache<int>();
  print(c.value);
  c.store(5);
  print(c.value);
}"#,
        ["null", "5"]
    };

    generic_static_factory_method => {
        r#"class Wrapper<T> {
  T data;
  Wrapper(this.data);
  static Wrapper<T> of<T>(T data) {
    return Wrapper(data);
  }
}
void main() {
  var w = Wrapper.of(42);
  print(w.data);
}"#,
        ["42"]
    };

    generic_result_ok_variant => {
        r#"class Result<T, E> {
  T? value;
  E? error;
  Result.ok(this.value) : error = null;
  Result.err(this.error) : value = null;
  bool get isOk {
    return error == null;
  }
}
void main() {
  var r = Result<int, String>.ok(10);
  print(r.isOk);
  print(r.value);
}"#,
        ["true", "10"]
    };

    generic_result_err_variant => {
        r#"class Result<T, E> {
  T? value;
  E? error;
  Result.ok(this.value) : error = null;
  Result.err(this.error) : value = null;
  bool get isOk {
    return error == null;
  }
}
void main() {
  var r = Result<int, String>.err('fail');
  print(r.isOk);
  print(r.error);
}"#,
        ["false", "fail"]
    };

    generic_function_with_callback => {
        r#"T apply<T>(T val, T Function(T) fn) {
  return fn(val);
}
void main() {
  print(apply(4, (n) => n * n));
}"#,
        ["16"]
    };

    iterable_typed_as_list_source => {
        r#"void main() {
  Iterable<int> it = [1, 2, 3];
  print(it.length);
  print(it.first);
}"#,
        ["3", "1"]
    };

    list_generic_from_method => {
        r#"List<T> singleton<T>(T value) {
  return [value];
}
void main() {
  var list = singleton('only');
  print(list.length);
  print(list.first);
}"#,
        ["1", "only"]
    };

    generic_queue_peek_and_dequeue => {
        r#"class Queue<T> {
  List<T> _data = [];
  void enqueue(T item) {
    _data.add(item);
  }
  T dequeue() {
    return _data.removeAt(0);
  }
  T peek() {
    return _data.first;
  }
}
void main() {
  var q = Queue<String>();
  q.enqueue('first');
  q.enqueue('second');
  print(q.peek());
  print(q.dequeue());
}"#,
        ["first", "first"]
    };

    generic_pair_equality_by_fields => {
        r#"class Pair<T> {
  T a;
  T b;
  Pair(this.a, this.b);
  bool sameFirst(Pair<T> other) {
    return a == other.a;
  }
}
void main() {
  var p1 = Pair(1, 2);
  var p2 = Pair(1, 9);
  print(p1.sameFirst(p2));
}"#,
        ["true"]
    };

    generic_map_values_list => {
        r#"Map<K, List<V>> bucket<K, V>(List<V> items, V Function(V) keyFn) {
  var map = <K, List<V>>{};
  for (var item in items) {
    var key = keyFn(item);
    map.putIfAbsent(key, () => []).add(item);
  }
  return map;
}
void main() {
  var m = bucket([1, 2, 3, 4], (n) => n % 2);
  print(m[0]!.length);
  print(m[1]!.length);
}"#,
        ["2", "2"]
    };

    typed_list_sort_in_place => {
        r#"void main() {
  List<int> nums = [3, 1, 4, 1];
  nums.sort();
  print(nums.join(','));
}"#,
        ["1,1,3,4"]
    };

    generic_optional_default => {
        r#"class Cell<T> {
  T? _val;
  T readOr(T fallback) {
    return _val ?? fallback;
  }
  void write(T v) {
    _val = v;
  }
}
void main() {
  var c = Cell<int>();
  print(c.readOr(-1));
  c.write(8);
  print(c.readOr(-1));
}"#,
        ["-1", "8"]
    };

    generic_three_type_params_triple => {
        r#"class Triple<A, B, C> {
  A first;
  B second;
  C third;
  Triple(this.first, this.second, this.third);
}
void main() {
  var t = Triple(1, 'x', true);
  print(t.first);
  print(t.second);
  print(t.third);
}"#,
        ["1", "x", "true"]
    };

    generic_method_returns_same_type => {
        r#"class Id<T> {
  T value;
  Id(this.value);
  Id<T> copy() {
    return Id(value);
  }
}
void main() {
  var a = Id(10);
  var b = a.copy();
  print(b.value);
}"#,
        ["10"]
    };

    map_generic_key_lookup => {
        r#"void main() {
  Map<int, String> table = {1: 'one', 2: 'two'};
  print(table[2]);
  print(table.containsKey(1));
}"#,
        ["two", "true"]
    };

    generic_list_first_where => {
        r#"T findFirst<T>(List<T> items, bool Function(T) test) {
  return items.firstWhere(test);
}
void main() {
  print(findFirst([1, 2, 3, 4], (n) => n > 2));
}"#,
        ["3"]
    };

    generic_enum_like_sealed_choice => {
        r#"class Choice<T> {
  T? some;
  bool get isSome {
    return some != null;
  }
  Choice.some(this.some);
  Choice.none() : some = null;
}
void main() {
  var c = Choice<int>.some(9);
  print(c.isSome);
  print(c.some);
}"#,
        ["true", "9"]
    };
}
