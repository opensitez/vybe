//! Dart cascade notation (`..`) on lists, maps, strings, StringBuffer, and custom objects.

dart_cases! {
    cascade_list_add_chain_reports_joined_elements => {
        r#"void main() {
  var nums = <int>[];
  nums..add(4)..add(5)..add(6);
  print(nums.join(','));
}"#,
        ["4,5,6"]
    };

    cascade_list_add_all_then_add_appends_tail => {
        r#"void main() {
  var nums = <int>[];
  nums..addAll([10, 20])..add(30);
  print(nums.join(','));
}"#,
        ["10,20,30"]
    };

    cascade_list_insert_then_add_preserves_order => {
        r#"void main() {
  var nums = [2, 3];
  nums..insert(0, 1)..add(4);
  print(nums.join(','));
}"#,
        ["1,2,3,4"]
    };

    cascade_list_remove_value_then_add_replaces_slot => {
        r#"void main() {
  var nums = [1, 2, 3];
  nums..remove(2)..add(9);
  print(nums.join(','));
}"#,
        ["1,3,9"]
    };

    cascade_list_remove_at_then_add_shifts_tail => {
        r#"void main() {
  var nums = [5, 6, 7];
  nums..removeAt(1)..add(8);
  print(nums.join(','));
}"#,
        ["5,7,8"]
    };

    cascade_list_remove_last_twice_shortens_length => {
        r#"void main() {
  var nums = [1, 2, 3, 4];
  nums..removeLast()..removeLast();
  print(nums.length);
  print(nums.join(','));
}"#,
        ["2", "1,2"]
    };

    cascade_list_clear_then_add_rebuilds_contents => {
        r#"void main() {
  var nums = [9, 9, 9];
  nums..clear()..add(1)..add(2);
  print(nums.isEmpty);
  print(nums.join(','));
}"#,
        ["false", "1,2"]
    };

    cascade_list_remove_range_then_add_truncates_middle => {
        r#"void main() {
  var nums = [1, 2, 3, 4, 5];
  nums..removeRange(1, 4)..add(6);
  print(nums.join(','));
}"#,
        ["1,5,6"]
    };

    cascade_list_expression_returns_original_receiver => {
        r#"void main() {
  var nums = <int>[];
  var same = nums..add(7)..add(8);
  print(same == nums);
  print(nums.length);
}"#,
        ["true", "2"]
    };

    cascade_list_add_then_sort_orders_ascending => {
        r#"void main() {
  var nums = [3, 1, 2];
  nums..add(4)..sort();
  print(nums.join(','));
}"#,
        ["1,2,3,4"]
    };

    cascade_map_add_chain_inserts_multiple_keys => {
        r#"void main() {
  var scores = <String, int>{};
  scores..add('alice', 10)..add('bob', 20)..add('carol', 30);
  print(scores['alice']);
  print(scores['bob']);
  print(scores['carol']);
}"#,
        ["10", "20", "30"]
    };

    cascade_map_add_all_then_add_merges_and_appends => {
        r#"void main() {
  var scores = <String, int>{'a': 1};
  scores..addAll({'b': 2, 'c': 3})..add('d', 4);
  print(scores.length);
  print(scores['d']);
}"#,
        ["4", "4"]
    };

    cascade_map_remove_then_add_replaces_entry => {
        r#"void main() {
  var scores = {'old': 1, 'keep': 2};
  scores..remove('old')..add('new', 99);
  print(scores.containsKey('old'));
  print(scores['new']);
}"#,
        ["false", "99"]
    };

    cascade_map_clear_then_add_starts_fresh => {
        r#"void main() {
  var scores = {'x': 1, 'y': 2};
  scores..clear()..add('only', 42);
  print(scores.length);
  print(scores['only']);
}"#,
        ["1", "42"]
    };

    cascade_map_add_overwrites_existing_key_via_chain => {
        r#"void main() {
  var scores = {'k': 5};
  scores..add('k', 9)..add('k', 11);
  print(scores['k']);
}"#,
        ["11"]
    };

    cascade_map_put_if_absent_chain_on_distinct_keys => {
        r#"void main() {
  var scores = <String, int>{};
  scores..putIfAbsent('first', () => 1)..putIfAbsent('second', () => 2);
  print(scores['first']);
  print(scores['second']);
}"#,
        ["1", "2"]
    };

    cascade_map_update_chain_mutates_values_in_place => {
        r#"void main() {
  var scores = {'n': 5};
  scores..update('n', (v) => v + 1)..update('n', (v) => v * 2);
  print(scores['n']);
}"#,
        ["12"]
    };

    cascade_map_expression_returns_original_receiver => {
        r#"void main() {
  var scores = <String, int>{};
  var same = scores..add('x', 1)..add('y', 2);
  print(same == scores);
  print(scores.length);
}"#,
        ["true", "2"]
    };

    cascade_string_upper_lower_chain_leaves_original_untouched => {
        r#"void main() {
  var word = 'dart';
  word..toUpperCase()..toLowerCase();
  print(word);
}"#,
        ["dart"]
    };

    cascade_string_pad_left_chain_preserves_original_value => {
        r#"void main() {
  var word = '7';
  word..padLeft(3, '0')..padLeft(5, '0');
  print(word);
}"#,
        ["7"]
    };

    cascade_string_trim_chain_keeps_surrounding_spaces => {
        r#"void main() {
  var word = '  hi  ';
  word..trim()..trimLeft();
  print(word);
}"#,
        ["  hi  "]
    };

    cascade_string_replace_first_chain_retains_source_text => {
        r#"void main() {
  var word = 'aaa';
  word..replaceFirst('a', 'b')..replaceFirst('a', 'c');
  print(word);
}"#,
        ["aaa"]
    };

    cascade_string_contains_and_starts_with_on_same_receiver => {
        r#"void main() {
  var word = 'cascade';
  word..contains('cas')..startsWith('cas');
  print(word.contains('cas'));
  print(word.startsWith('cas'));
}"#,
        ["true", "true"]
    };

    cascade_string_index_methods_chain_on_same_receiver => {
        r#"void main() {
  var word = 'ababa';
  word..indexOf('b')..lastIndexOf('b');
  print(word.indexOf('b'));
  print(word.lastIndexOf('b'));
}"#,
        ["1", "3"]
    };

    cascade_string_substring_chain_preserves_full_length => {
        r#"void main() {
  var word = 'abcdef';
  word..substring(1, 3)..substring(4);
  print(word.length);
  print(word);
}"#,
        ["6", "abcdef"]
    };

    cascade_string_split_chain_keeps_original_literal => {
        r#"void main() {
  var csv = 'a,b,c';
  csv..split(',')..split('b');
  print(csv);
}"#,
        ["a,b,c"]
    };

    cascade_string_buffer_write_twice_builds_concatenation => {
        r#"void main() {
  var buf = StringBuffer();
  buf..write('foo')..write('bar');
  print(buf.toString());
}"#,
        ["foobar"]
    };

    cascade_string_buffer_write_then_writeln_adds_line_break => {
        r#"void main() {
  var buf = StringBuffer();
  buf..write('line')..writeln('end');
  print(buf.toString());
}"#,
        ["lineend\n"]
    };

    cascade_string_buffer_write_chain_builds_separated_tokens => {
        r#"void main() {
  var buf = StringBuffer();
  buf..write('x')..write('-')..write('y')..write('-')..write('z')..write('!');
  print(buf.toString());
}"#,
        ["x-y-z!"]
    };

    cascade_string_buffer_clear_after_writes_empties_buffer => {
        r#"void main() {
  var buf = StringBuffer();
  buf..write('gone')..clear()..write('ok');
  print(buf.toString());
}"#,
        ["ok"]
    };

    cascade_string_buffer_expression_returns_original_receiver => {
        r#"void main() {
  var buf = StringBuffer();
  var same = buf..write('a')..write('b');
  print(same == buf);
  print(buf.toString());
}"#,
        ["true", "ab"]
    };

    cascade_custom_counter_increment_chain_accumulates => {
        r#"class Counter {
  int value = 0;
  void bump() { value += 1; }
}
void main() {
  var tally = Counter();
  tally..bump()..bump()..bump();
  print(tally.value);
}"#,
        ["3"]
    };

    cascade_custom_setter_chain_updates_backing_field => {
        r#"class Box {
  int _size = 0;
  set size(int v) { _size = v; }
  int get size => _size;
}
void main() {
  var crate = Box();
  crate..size = 10..size = 25;
  print(crate.size);
}"#,
        ["25"]
    };

    cascade_custom_method_then_setter_combines_mutations => {
        r#"class Widget {
  String label = '';
  void setLabel(String s) { label = s; }
  set prefix(String p) { label = p + label; }
}
void main() {
  var w = Widget();
  w..setLabel('base')..prefix = 'pre-';
  print(w.label);
}"#,
        ["pre-base"]
    };

    cascade_custom_nested_list_field_mutated_via_chain => {
        r#"class Bag {
  List<int> items = [];
  void seed(int n) { items.add(n); }
}
void main() {
  var bag = Bag();
  bag..seed(1)..items.add(2)..items.add(3);
  print(bag.items.join(','));
}"#,
        ["1,2,3"]
    };

    cascade_custom_builder_style_chain_configures_object => {
        r#"class Report {
  String title = '';
  int pages = 0;
  void setTitle(String t) { title = t; }
  void addPage() { pages += 1; }
}
void main() {
  var doc = Report();
  doc..setTitle('Q1')..addPage()..addPage();
  print(doc.title);
  print(doc.pages);
}"#,
        ["Q1", "2"]
    };
}
