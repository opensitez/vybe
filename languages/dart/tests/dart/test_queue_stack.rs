//! Queue FIFO via dart:collection and List-based LIFO stack patterns.

dart_cases! {
    queue_empty_starts_with_zero_length => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  print(q.length);
}"#,
        ["0"]
    };

    queue_add_then_remove_first_returns_value => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(10);
  print(q.removeFirst());
}"#,
        ["10"]
    };

    queue_fifo_order_two_items => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<String>();
  q.add('first');
  q.add('second');
  print(q.removeFirst());
  print(q.removeFirst());
}"#,
        ["first", "second"]
    };

    queue_fifo_order_three_items => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  q.add(2);
  q.add(3);
  print(q.removeFirst());
  print(q.removeFirst());
  print(q.removeFirst());
}"#,
        ["1", "2", "3"]
    };

    queue_length_tracks_enqueued_items => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  q.add(2);
  print(q.length);
}"#,
        ["2"]
    };

    queue_is_empty_before_add => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  print(q.isEmpty);
}"#,
        ["true"]
    };

    queue_is_not_empty_after_add => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  print(q.isNotEmpty);
}"#,
        ["true"]
    };

    queue_first_peeks_front_element => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(5);
  q.add(6);
  print(q.first);
}"#,
        ["5"]
    };

    queue_last_peeks_back_element => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(5);
  q.add(6);
  print(q.last);
}"#,
        ["6"]
    };

    queue_add_after_remove_first_continues_fifo => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  q.removeFirst();
  q.add(2);
  print(q.removeFirst());
}"#,
        ["2"]
    };

    queue_add_all_preserves_order => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.addAll([1, 2, 3]);
  print(q.removeFirst());
  print(q.removeFirst());
  print(q.removeFirst());
}"#,
        ["1", "2", "3"]
    };

    queue_clear_empties_elements => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  q.add(2);
  q.clear();
  print(q.isEmpty);
  print(q.length);
}"#,
        ["true", "0"]
    };

    queue_remove_last_from_back => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  q.add(2);
  q.add(3);
  print(q.removeLast());
  print(q.first);
}"#,
        ["3", "1"]
    };

    queue_contains_existing_value => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<String>();
  q.add('a');
  q.add('b');
  print(q.contains('b'));
}"#,
        ["true"]
    };

    queue_contains_missing_value => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  print(q.contains(9));
}"#,
        ["false"]
    };

    queue_to_list_preserves_fifo_order => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(4);
  q.add(5);
  q.add(6);
  print(q.toList().join(','));
}"#,
        ["4,5,6"]
    };

    queue_from_constructor_with_initial_items => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>.from([7, 8]);
  print(q.removeFirst());
  print(q.removeFirst());
}"#,
        ["7", "8"]
    };

    queue_remove_until_empty => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  q.add(2);
  q.removeFirst();
  q.removeFirst();
  print(q.isEmpty);
}"#,
        ["true"]
    };

    queue_for_in_iteration_fifo_order => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  q.add(2);
  q.add(3);
  var seen = <int>[];
  for (var item in q) {
    seen.add(item);
  }
  print(seen.join(','));
}"#,
        ["1,2,3"]
    };

    queue_mixed_add_and_remove_first_sequence => {
        r#"import 'dart:collection';
void main() {
  var q = Queue<int>();
  q.add(1);
  q.add(2);
  print(q.removeFirst());
  q.add(3);
  print(q.removeFirst());
  print(q.removeFirst());
}"#,
        ["1", "2", "3"]
    };

    list_stack_empty_starts_with_zero_length => {
        r#"void main() {
  var stack = <int>[];
  print(stack.length);
}"#,
        ["0"]
    };

    list_stack_add_then_remove_last_returns_top => {
        r#"void main() {
  var stack = <int>[];
  stack.add(10);
  print(stack.removeLast());
}"#,
        ["10"]
    };

    list_stack_lifo_two_items => {
        r#"void main() {
  var stack = <String>[];
  stack.add('bottom');
  stack.add('top');
  print(stack.removeLast());
  print(stack.removeLast());
}"#,
        ["top", "bottom"]
    };

    list_stack_lifo_three_items => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  stack.add(2);
  stack.add(3);
  print(stack.removeLast());
  print(stack.removeLast());
  print(stack.removeLast());
}"#,
        ["3", "2", "1"]
    };

    list_stack_length_tracks_pushes => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  stack.add(2);
  print(stack.length);
}"#,
        ["2"]
    };

    list_stack_is_empty_before_push => {
        r#"void main() {
  var stack = <int>[];
  print(stack.isEmpty);
}"#,
        ["true"]
    };

    list_stack_is_not_empty_after_push => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  print(stack.isNotEmpty);
}"#,
        ["true"]
    };

    list_stack_last_peeks_top_element => {
        r#"void main() {
  var stack = <int>[];
  stack.add(5);
  stack.add(6);
  print(stack.last);
}"#,
        ["6"]
    };

    list_stack_first_is_bottom_element => {
        r#"void main() {
  var stack = <int>[];
  stack.add(5);
  stack.add(6);
  print(stack.first);
}"#,
        ["5"]
    };

    list_stack_push_after_pop_continues_lifo => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  stack.removeLast();
  stack.add(2);
  print(stack.removeLast());
}"#,
        ["2"]
    };

    list_stack_add_all_then_pop_sequence => {
        r#"void main() {
  var stack = <int>[];
  stack.addAll([1, 2, 3]);
  print(stack.removeLast());
  print(stack.removeLast());
  print(stack.removeLast());
}"#,
        ["3", "2", "1"]
    };

    list_stack_clear_empties_elements => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  stack.add(2);
  stack.clear();
  print(stack.isEmpty);
  print(stack.length);
}"#,
        ["true", "0"]
    };

    list_stack_remove_last_shortens_length => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  stack.add(2);
  stack.removeLast();
  print(stack.length);
  print(stack.last);
}"#,
        ["1", "1"]
    };

    list_stack_pop_until_empty => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  stack.add(2);
  stack.removeLast();
  stack.removeLast();
  print(stack.isEmpty);
}"#,
        ["true"]
    };

    list_stack_duplicate_values_lifo_order => {
        r#"void main() {
  var stack = <int>[];
  stack.add(7);
  stack.add(7);
  print(stack.removeLast());
  print(stack.removeLast());
}"#,
        ["7", "7"]
    };

    list_stack_typed_int_lifo_join => {
        r#"void main() {
  List<int> stack = [];
  stack.add(1);
  stack.add(2);
  stack.add(3);
  var out = <int>[];
  out.add(stack.removeLast());
  out.add(stack.removeLast());
  out.add(stack.removeLast());
  print(out.join('-'));
}"#,
        ["3-2-1"]
    };

    list_stack_mixed_push_pop_pattern => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  stack.add(2);
  print(stack.removeLast());
  stack.add(3);
  print(stack.removeLast());
  print(stack.removeLast());
}"#,
        ["2", "3", "1"]
    };

    list_stack_for_in_bottom_to_top => {
        r#"void main() {
  var stack = <int>[];
  stack.add(1);
  stack.add(2);
  stack.add(3);
  print(stack.join(','));
}"#,
        ["1,2,3"]
    };

    list_stack_contains_top_value => {
        r#"void main() {
  var stack = <String>[];
  stack.add('a');
  stack.add('b');
  print(stack.contains('b'));
}"#,
        ["true"]
    };

    list_stack_sublist_copy_does_not_change_stack => {
        r#"void main() {
  var stack = <int>[1, 2, 3];
  var copy = stack.sublist(0);
  copy.add(4);
  print(stack.length);
  print(stack.last);
}"#,
        ["3", "3"]
    };
}
