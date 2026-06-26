//! Deep augmented assignment: += on map[key], list[i], strings, cascades, and getter/setter targets.

dart_cases! {
    map_bracket_key_plus_assign_int => {
        r#"void main() {
  var m = {'a': 1};
  m['a'] += 2;
  print(m['a']);
}"#,
        ["3"]
    };

    map_bracket_key_plus_assign_from_zero => {
        r#"void main() {
  var m = <String, int>{'k': 0};
  m['k'] += 5;
  print(m['k']);
}"#,
        ["5"]
    };

    map_dynamic_key_plus_assign => {
        r#"void main() {
  var m = {'x': 10};
  var key = 'x';
  m[key] += 3;
  print(m[key]);
}"#,
        ["13"]
    };

    map_multiple_keys_plus_assign => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  m['a'] += 4;
  m['b'] += 6;
  print(m['a']);
  print(m['b']);
}"#,
        ["5", "8"]
    };

    list_index_plus_assign => {
        r#"void main() {
  var xs = [1, 2, 3];
  xs[1] += 5;
  print(xs[1]);
}"#,
        ["7"]
    };

    list_first_index_plus_assign => {
        r#"void main() {
  var xs = [10, 20];
  xs[0] += 1;
  print(xs[0]);
}"#,
        ["11"]
    };

    list_last_index_plus_assign => {
        r#"void main() {
  var xs = [10, 20, 30];
  xs[2] += 7;
  print(xs[2]);
}"#,
        ["37"]
    };

    list_index_via_variable_plus_assign => {
        r#"void main() {
  var xs = [5, 6, 7];
  var i = 1;
  xs[i] += 2;
  print(xs[i]);
}"#,
        ["8"]
    };

    string_plus_assign_append => {
        r#"void main() {
  var s = 'hello';
  s += ' world';
  print(s);
}"#,
        ["hello world"]
    };

    string_plus_assign_char => {
        r#"void main() {
  var s = 'ab';
  s += 'c';
  print(s);
}"#,
        ["abc"]
    };

    string_plus_assign_empty_suffix => {
        r#"void main() {
  var s = 'dart';
  s += '';
  print(s.length);
}"#,
        ["4"]
    };

    cascade_then_field_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var c = Counter();
  c..value += 3..value += 2;
  print(c.value);
}"#,
        ["5"]
    };

    cascade_list_add_then_index_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var box = Box();
  box..items.add(1)..items[0] += 4;
  print(box.items[0]);
}"#,
        ["5"]
    };

    getter_setter_instance_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var acc = Accumulator();
  acc.total += 10;
  acc.total += 5;
  print(acc.total);
}"#,
        ["15"]
    };

    getter_setter_static_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  Score.points += 3;
  Score.points += 2;
  print(Score.points);
}"#,
        ["5"]
    };

    map_value_minus_assign => {
        r#"void main() {
  var m = {'n': 20};
  m['n'] -= 5;
  print(m['n']);
}"#,
        ["15"]
    };

    list_index_minus_assign => {
        r#"void main() {
  var xs = [100];
  xs[0] -= 25;
  print(xs[0]);
}"#,
        ["75"]
    };

    map_value_times_assign => {
        r#"void main() {
  var m = {'n': 3};
  m['n'] *= 4;
  print(m['n']);
}"#,
        ["12"]
    };

    list_index_times_assign => {
        r#"void main() {
  var xs = [2, 3];
  xs[0] *= 5;
  print(xs[0]);
}"#,
        ["10"]
    };

    map_value_divide_assign => {
        r#"void main() {
  var m = {'n': 20};
  m['n'] /= 4;
  print(m['n']);
}"#,
        ["5.0"]
    };

    list_index_integer_divide_assign => {
        r#"void main() {
  var xs = [10];
  xs[0] ~/= 3;
  print(xs[0]);
}"#,
        ["3"]
    };

    map_value_modulo_assign => {
        r#"void main() {
  var m = {'n': 10};
  m['n'] %= 3;
  print(m['n']);
}"#,
        ["1"]
    };

    string_multiple_plus_assign => {
        r#"void main() {
  var s = 'a';
  s += 'b';
  s += 'c';
  print(s);
}"#,
        ["abc"]
    };

    cascade_map_entry_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var holder = Holder();
  holder..data['x'] += 2..data['x'] += 3;
  print(holder.data['x']);
}"#,
        ["5"]
    };

    cascade_on_list_then_index_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var w = Wrapper();
  w..nums[0] += 1..nums[1] += 2;
  print(w.nums[0]);
  print(w.nums[1]);
}"#,
        ["1", "2"]
    };

    getter_with_validation_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var g = Guarded();
  g.level += 2;
  g.level += 3;
  print(g.level);
}"#,
        ["5"]
    };

    local_var_plus_assign => {
        r#"void main() {
  var n = 1;
  n += 9;
  print(n);
}"#,
        ["10"]
    };

    local_double_plus_assign => {
        r#"void main() {
  var d = 1.5;
  d += 2.5;
  print(d);
}"#,
        ["4.0"]
    };

    map_nested_key_sequence_plus_assign => {
        r#"void main() {
  var m = {'a': 1, 'b': 1};
  m['a'] += 1;
  m['b'] += 2;
  print(m['a']! + m['b']!);
}"#,
        ["5"]
    };

    list_two_indices_plus_assign_same_pass => {
        r#"void main() {
  var xs = [1, 1, 1];
  xs[0] += 2;
  xs[2] += 3;
  print(xs[0]);
  print(xs[2]);
}"#,
        ["3", "4"]
    };

    string_plus_assign_number_to_string => {
        r#"void main() {
  var s = 'n=';
  s += 42.toString();
  print(s);
}"#,
        ["n=42"]
    };

    cascade_string_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var b = Buffer('x');
  b..text += 'y'..text += 'z';
  print(b.text);
}"#,
        ["xyz"]
    };

    instance_field_plus_assign_after_read => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var p = Pair();
  p.a += p.b;
  print(p.a);
}"#,
        ["5"]
    };

    map_bracket_bitwise_or_assign => {
        r#"void main() {
  var m = {'f': 1};
  m['f'] |= 2;
  print(m['f']);
}"#,
        ["3"]
    };

    list_index_bitwise_and_assign => {
        r#"void main() {
  var xs = [15];
  xs[0] &= 10;
  print(xs[0]);
}"#,
        ["10"]
    };

    map_bracket_bitwise_xor_assign => {
        r#"void main() {
  var m = {'f': 5};
  m['f'] ^= 1;
  print(m['f']);
}"#,
        ["4"]
    };

    getter_only_side_effect_via_plus_assign => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var t = Tally();
  t.count += 1;
  t.count += 1;
  print(t.count);
}"#,
        ["2"]
    };

    cascade_then_string_plus_assign_on_field => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var n = Note('hi');
  n..body += '!'..body += '?';
  print(n.body);
}"#,
        ["hi!?"]
    };

    list_growable_middle_index_plus_assign => {
        r#"void main() {
  var xs = [0, 0, 0, 0, 0];
  xs[2] += 9;
  print(xs[2]);
}"#,
        ["9"]
    };

    map_string_key_plus_assign_twice => {
        r#"void main() {
  var m = {'msg': 'a'};
  m['msg'] += 'b';
  m['msg'] += 'c';
  print(m['msg']);
}"#,
        ["abc"]
    };

    list_index_plus_assign_preserves_length => {
        r#"void main() {
  var xs = [1, 2, 3];
  xs[1] += 10;
  print(xs.length);
  print(xs[1]);
}"#,
        ["3", "12"]
    };

    cascade_new_object_then_plus_assign_field => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var v = (Counter()..value += 4).value;
  print(v);
}"#,
        ["4"]
    };

    map_plus_assign_then_read_other_key => {
        r#"void main() {
  var m = {'a': 1, 'b': 10};
  m['a'] += 2;
  print(m['b']);
  print(m['a']);
}"#,
        ["10", "3"]
    };

    string_plus_assign_in_loop_build => {
        r#"void main() {
  var s = '';
  for (var i = 0; i < 3; i++) {
    s += i.toString();
  }
  print(s);
}"#,
        ["012"]
    };

    getter_setter_plus_assign_returns_updated => {
        r#"class Counter { int value = 0; }
class Box { List<int> items = [0]; }
class Accumulator { int _t = 0; int get total => _t; set total(int v) { _t = v; } }
class Score { static int points = 0; }
class Holder { Map<String, int> data = {'x': 0}; }
class Wrapper { List<int> nums = [0, 0]; }
class Guarded { int _level = 0; int get level => _level; set level(int v) { _level = v; } }
class Pair { int a = 2; int b = 3; }
class Tally { int _c = 0; int get count => _c; set count(int v) { _c = v; } }
class Buffer { String text; Buffer(this.text); }
class Note { String body; Note(this.body); }
class Wallet { int _b = 0; int get balance => _b; set balance(int v) { _b = v; } }
void main() {
  var w = Wallet();
  w.balance += 100;
  w.balance -= 30;
  print(w.balance);
}"#,
        ["70"]
    };
}
