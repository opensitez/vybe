//! Generic type inference: call-site inference, var with generic returns,
//! argument-driven inference, and bounded type parameters with extends.

dart_cases! {
    infer_generic_identity_from_int_literal => {
        r#"T id<T>(T v) => v;
void main() {
  print(id(42));
}"#,
        ["42"]
    };

    infer_generic_identity_from_string_literal => {
        r#"T id<T>(T v) => v;
void main() {
  print(id('dart'));
}"#,
        ["dart"]
    };

    infer_generic_identity_from_bool_literal => {
        r#"T id<T>(T v) => v;
void main() {
  print(id(true));
}"#,
        ["true"]
    };

    infer_generic_identity_from_double_literal => {
        r#"T id<T>(T v) => v;
void main() {
  print(id(3.5));
}"#,
        ["3.5"]
    };

    infer_pair_types_from_mixed_arguments => {
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

    infer_pair_types_reversed_argument_order => {
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

    infer_box_type_from_constructor_argument => {
        r#"class Box<T> {
  T value;
  Box(this.value);
}
void main() {
  var b = Box(100);
  print(b.value);
}"#,
        ["100"]
    };

    infer_box_type_from_string_constructor_arg => {
        r#"class Box<T> {
  T value;
  Box(this.value);
}
void main() {
  var b = Box('hello');
  print(b.value);
}"#,
        ["hello"]
    };

    infer_list_element_type_in_var_assignment => {
        r#"void main() {
  var nums = [1, 2, 3];
  print(nums.length);
  print(nums[1]);
}"#,
        ["3", "2"]
    };

    infer_list_element_type_string_items => {
        r#"void main() {
  var words = ['a', 'b', 'c'];
  print(words.join('-'));
}"#,
        ["a-b-c"]
    };

    infer_map_key_value_types_from_literals => {
        r#"void main() {
  var scores = {'Ada': 90, 'Bob': 85};
  print(scores['Ada']);
  print(scores.length);
}"#,
        ["90", "2"]
    };

    infer_set_element_type_from_literals => {
        r#"void main() {
  var tags = {'dart', 'flutter', 'dart'};
  print(tags.length);
  print(tags.contains('dart'));
}"#,
        ["2", "true"]
    };

    infer_generic_function_return_via_var => {
        r#"List<T> singleton<T>(T value) {
  return [value];
}
void main() {
  var list = singleton(7);
  print(list.length);
  print(list.first);
}"#,
        ["1", "7"]
    };

    infer_generic_function_return_string_via_var => {
        r#"List<T> singleton<T>(T value) {
  return [value];
}
void main() {
  var list = singleton('only');
  print(list.first);
}"#,
        ["only"]
    };

    infer_repeat_list_from_int_argument => {
        r#"List<T> repeat<T>(T value, int times) {
  return List.generate(times, (_) => value);
}
void main() {
  var list = repeat(5, 3);
  print(list.join(','));
}"#,
        ["5,5,5"]
    };

    infer_repeat_list_from_string_argument => {
        r#"List<T> repeat<T>(T value, int times) {
  return List.generate(times, (_) => value);
}
void main() {
  var list = repeat('x', 2);
  print(list.join(''));
}"#,
        ["xx"]
    };

    infer_first_of_list_from_element_type => {
        r#"T firstOf<T>(List<T> items) {
  return items.first;
}
void main() {
  print(firstOf([10, 20, 30]));
}"#,
        ["10"]
    };

    infer_first_of_list_string_elements => {
        r#"T firstOf<T>(List<T> items) {
  return items.first;
}
void main() {
  print(firstOf(['alpha', 'beta']));
}"#,
        ["alpha"]
    };

    infer_last_of_list_from_element_type => {
        r#"T lastOf<T>(List<T> items) {
  return items.last;
}
void main() {
  print(lastOf([1, 2, 3]));
}"#,
        ["3"]
    };

    infer_map_get_inferred_value_type => {
        r#"V? lookup<K, V>(Map<K, V> map, K key) {
  return map[key];
}
void main() {
  var table = {'a': 1, 'b': 2};
  print(lookup(table, 'b'));
}"#,
        ["2"]
    };

    infer_bounded_add_nums_from_int_args => {
        r#"T addNums<T extends num>(T a, T b) {
  return (a + b) as T;
}
void main() {
  print(addNums(3, 4));
}"#,
        ["7"]
    };

    infer_bounded_add_nums_from_double_args => {
        r#"T addNums<T extends num>(T a, T b) {
  return (a + b) as T;
}
void main() {
  print(addNums(1.5, 2.5));
}"#,
        ["4.0"]
    };

    infer_comparable_max_from_int_args => {
        r#"T maxOf<T extends Comparable<T>>(T a, T b) {
  return a.compareTo(b) >= 0 ? a : b;
}
void main() {
  print(maxOf(3, 9));
}"#,
        ["9"]
    };

    infer_comparable_max_from_string_args => {
        r#"T maxOf<T extends Comparable<T>>(T a, T b) {
  return a.compareTo(b) >= 0 ? a : b;
}
void main() {
  print(maxOf('apple', 'banana'));
}"#,
        ["banana"]
    };

    infer_comparable_min_from_string_args => {
        r#"T minOf<T extends Comparable<T>>(T a, T b) {
  return a.compareTo(b) <= 0 ? a : b;
}
void main() {
  print(minOf('zebra', 'ant'));
}"#,
        ["ant"]
    };

    infer_min_of_ints_via_bounded_comparable => {
        r#"T minOf<T extends Comparable<T>>(T a, T b) {
  return a.compareTo(b) <= 0 ? a : b;
}
void main() {
  print(minOf(10, 4));
}"#,
        ["4"]
    };

    infer_generic_method_map_all_on_int_list => {
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

    infer_generic_method_map_all_to_string => {
        r#"class Holder<T> {
  List<T> items;
  Holder(this.items);
  List<R> mapAll<R>(R Function(T) fn) {
    return items.map(fn).toList();
  }
}
void main() {
  var h = Holder([1, 2, 3]);
  var labels = h.mapAll((n) => 'n$n');
  print(labels.join('|'));
}"#,
        ["n1|n2|n3"]
    };

    infer_generic_method_keep_where_on_ints => {
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

    infer_generic_stack_push_pop_int => {
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
  var s = Stack();
  s.push(1);
  s.push(2);
  print(s.pop());
}"#,
        ["2"]
    };

    infer_generic_stack_push_pop_string => {
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
  var s = Stack();
  s.push('a');
  s.push('b');
  print(s.pop());
}"#,
        ["b"]
    };

    infer_generic_queue_enqueue_dequeue => {
        r#"class Queue<T> {
  List<T> _data = [];
  void enqueue(T item) {
    _data.add(item);
  }
  T dequeue() {
    return _data.removeAt(0);
  }
}
void main() {
  var q = Queue();
  q.enqueue('first');
  q.enqueue('second');
  print(q.dequeue());
}"#,
        ["first"]
    };

    infer_swap_pair_from_argument_types => {
        r#"class Pair<T, U> {
  T first;
  U second;
  Pair(this.first, this.second);
}
Pair<U, T> swap<T, U>(Pair<T, U> p) {
  return Pair(p.second, p.first);
}
void main() {
  var p = Pair(1, 'a');
  var s = swap(p);
  print(s.first);
  print(s.second);
}"#,
        ["a", "1"]
    };

    infer_apply_with_int_callback => {
        r#"T apply<T>(T val, T Function(T) fn) {
  return fn(val);
}
void main() {
  print(apply(4, (n) => n * n));
}"#,
        ["16"]
    };

    infer_apply_with_string_callback => {
        r#"T apply<T>(T val, T Function(T) fn) {
  return fn(val);
}
void main() {
  print(apply('hi', (s) => s + s));
}"#,
        ["hihi"]
    };

    infer_fold_accumulator_type_from_seed => {
        r#"T foldList<T>(List<T> items, T Function(T, T) combine) {
  var acc = items.first;
  for (var i = 1; i < items.length; i++) {
    acc = combine(acc, items[i]);
  }
  return acc;
}
void main() {
  print(foldList([1, 2, 3], (a, b) => a + b));
}"#,
        ["6"]
    };

    infer_fold_string_concat_from_seed => {
        r#"T foldList<T>(List<T> items, T Function(T, T) combine) {
  var acc = items.first;
  for (var i = 1; i < items.length; i++) {
    acc = combine(acc, items[i]);
  }
  return acc;
}
void main() {
  print(foldList(['a', 'b'], (a, b) => a + b));
}"#,
        ["ab"]
    };

    infer_optional_of_from_non_null_value => {
        r#"T? wrap<T>(T? value) {
  return value;
}
void main() {
  var v = wrap(99);
  print(v);
}"#,
        ["99"]
    };

    infer_optional_of_null_for_int => {
        r#"T? wrap<T>(T? value) {
  return value;
}
void main() {
  var v = wrap(null);
  print(v);
}"#,
        ["null"]
    };

    infer_triple_three_distinct_types => {
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

    infer_static_factory_of_from_argument => {
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

    infer_static_factory_of_string => {
        r#"class Wrapper<T> {
  T data;
  Wrapper(this.data);
  static Wrapper<T> of<T>(T data) {
    return Wrapper(data);
  }
}
void main() {
  var w = Wrapper.of('ok');
  print(w.data);
}"#,
        ["ok"]
    };

    infer_list_from_empty_literal_with_context => {
        r#"List<int> build() {
  return [];
}
void main() {
  var list = build();
  print(list.length);
  print(list.isEmpty);
}"#,
        ["0", "true"]
    };

    infer_map_values_from_nested_list_literal => {
        r#"Map<int, List<String>> build() {
  return {1: ['a'], 2: ['b', 'c']};
}
void main() {
  var m = build();
  print(m[2]!.length);
  print(m[2]![0]);
}"#,
        ["2", "b"]
    };

    infer_find_first_predicate_on_int_list => {
        r#"T findFirst<T>(List<T> items, bool Function(T) test) {
  return items.firstWhere(test);
}
void main() {
  print(findFirst([1, 2, 3, 4], (n) => n > 2));
}"#,
        ["3"]
    };

    infer_zip_two_lists_of_same_inferred_type => {
        r#"List<Pair<T, T>> zipSame<T>(List<T> a, List<T> b) {
  var out = <Pair<T, T>>[];
  for (var i = 0; i < a.length && i < b.length; i++) {
    out.add(Pair(a[i], b[i]));
  }
  return out;
}
class Pair<A, B> {
  A first;
  B second;
  Pair(this.first, this.second);
}
void main() {
  var z = zipSame([1, 2], [3, 4]);
  print(z[0].first);
  print(z[1].second);
}"#,
        ["1", "4"]
    };

    infer_cast_list_after_generic_map => {
        r#"List<R> convert<T, R>(List<T> items, R Function(T) fn) {
  return items.map(fn).toList();
}
void main() {
  var lengths = convert(['ab', 'cde'], (s) => s.length);
  print(lengths.join(','));
}"#,
        ["2,3"]
    };

    infer_bounded_object_container_from_int => {
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

    infer_bounded_object_container_from_string => {
        r#"class Container<T extends Object> {
  T value;
  Container(this.value);
}
void main() {
  var c = Container('data');
  print(c.value);
}"#,
        ["data"]
    };

    infer_nested_generic_call_chain => {
        r#"T id<T>(T v) => v;
List<T> singleton<T>(T v) => [id(v)];
void main() {
  var list = singleton(8);
  print(list.first);
}"#,
        ["8"]
    };

    infer_var_type_from_generic_method_return => {
        r#"class Cell<T> {
  T value;
  Cell(this.value);
  Cell<R> map<R>(R Function(T) fn) {
    return Cell(fn(value));
  }
}
void main() {
  var c = Cell(5);
  var mapped = c.map((n) => n.toString());
  print(mapped.value);
}"#,
        ["5"]
    };

    infer_list_generate_with_inferred_element => {
        r#"List<T> duplicate<T>(T value, int count) {
  return List.generate(count, (_) => value);
}
void main() {
  var items = duplicate(true, 2);
  print(items[0]);
  print(items[1]);
}"#,
        ["true", "true"]
    };
}
