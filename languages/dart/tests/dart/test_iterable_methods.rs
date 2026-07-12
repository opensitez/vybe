//! Iterable higher-order methods: fold, reduce, expand, followedBy, skipWhile,
//! takeWhile, cast, toSet, toList, singleWhere, elementAt, and related edge cases.

dart_cases! {
    iterable_fold_with_seed_sums_list => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3, 4];
  print(nums.fold(10, (acc, n) => acc + n));
}"#,
        ["20"]
    };

    iterable_fold_string_concatenation => {
        r#"void main() {
  Iterable<String> words = ['a', 'b', 'c'];
  print(words.fold('', (acc, w) => acc + w));
}"#,
        ["abc"]
    };

    iterable_fold_on_empty_with_seed => {
        r#"void main() {
  Iterable<int> empty = [];
  print(empty.fold(5, (acc, n) => acc + n));
}"#,
        ["5"]
    };

    iterable_reduce_multiplies_elements => {
        r#"void main() {
  Iterable<int> nums = [2, 3, 4];
  print(nums.reduce((a, b) => a * b));
}"#,
        ["24"]
    };

    iterable_reduce_concatenates_strings => {
        r#"void main() {
  Iterable<String> parts = ['x', 'y', 'z'];
  print(parts.reduce((a, b) => a + b));
}"#,
        ["xyz"]
    };

    iterable_reduce_on_map_values => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  print(m.values.reduce((a, b) => a + b));
}"#,
        ["6"]
    };

    iterable_expand_flattens_nested_lists => {
        r#"void main() {
  Iterable<List<int>> nested = [[1, 2], [3], [4, 5]];
  print(nested.expand((part) => part).join(','));
}"#,
        ["1,2,3,4,5"]
    };

    iterable_expand_empty_inner_lists => {
        r#"void main() {
  Iterable<List<int>> nested = [[], [1], [], [2, 3]];
  print(nested.expand((part) => part).join(','));
}"#,
        ["1,2,3"]
    };

    iterable_expand_from_string_chars => {
        r#"void main() {
  Iterable<String> words = ['ab', 'c'];
  print(words.expand((w) => w.split('')).join(''));
}"#,
        ["abc"]
    };

    iterable_followed_by_concatenates_two_lists => {
        r#"void main() {
  Iterable<int> first = [1, 2];
  var combined = first.followedBy([3, 4]);
  print(combined.join(','));
}"#,
        ["1,2,3,4"]
    };

    iterable_followed_by_with_empty_suffix => {
        r#"void main() {
  Iterable<int> first = [7, 8];
  print(first.followedBy([]).join(','));
}"#,
        ["7,8"]
    };

    iterable_followed_by_empty_prefix => {
        r#"void main() {
  Iterable<int> empty = [];
  print(empty.followedBy([1, 2]).join(','));
}"#,
        ["1,2"]
    };

    iterable_skip_while_drops_leading_matches => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3, 4, 1];
  print(nums.skipWhile((n) => n < 3).join(','));
}"#,
        ["3,4,1"]
    };

    iterable_skip_while_on_all_matching => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3];
  print(nums.skipWhile((n) => n < 10).length);
}"#,
        ["0"]
    };

    iterable_skip_while_on_none_matching => {
        r#"void main() {
  Iterable<int> nums = [5, 6, 7];
  print(nums.skipWhile((n) => n < 0).join(','));
}"#,
        ["5,6,7"]
    };

    iterable_take_while_keeps_leading_matches => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3, 4, 1];
  print(nums.takeWhile((n) => n < 3).join(','));
}"#,
        ["1,2"]
    };

    iterable_take_while_on_all_matching => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3];
  print(nums.takeWhile((n) => n < 10).join(','));
}"#,
        ["1,2,3"]
    };

    iterable_take_while_on_none_matching => {
        r#"void main() {
  Iterable<int> nums = [5, 6, 7];
  print(nums.takeWhile((n) => n < 0).length);
}"#,
        ["0"]
    };

    iterable_cast_narrows_object_list => {
        r#"void main() {
  Iterable<Object> objs = [1, 2, 3];
  var nums = objs.cast<int>();
  print(nums.reduce((a, b) => a + b));
}"#,
        ["6"]
    };

    iterable_cast_preserves_order => {
        r#"void main() {
  Iterable<Object> objs = ['a', 'b', 'c'];
  var strs = objs.cast<String>();
  print(strs.join('-'));
}"#,
        ["a-b-c"]
    };

    iterable_to_set_deduplicates_elements => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 2, 3, 3, 3];
  var s = nums.toSet();
  print(s.length);
  print(s.contains(2));
}"#,
        ["3", "true"]
    };

    iterable_to_set_from_map_keys => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  var keys = m.keys.toSet();
  print(keys.length);
  print(keys.contains('b'));
}"#,
        ["3", "true"]
    };

    iterable_to_list_materializes_lazy_chain => {
        r#"void main() {
  Iterable<int> it = [1, 2, 3].map((n) => n * 2);
  var list = it.toList();
  print(list.join(','));
}"#,
        ["2,4,6"]
    };

    iterable_to_list_from_set => {
        r#"void main() {
  Iterable<int> it = {3, 1, 2};
  var list = it.toList();
  list.sort();
  print(list.join(','));
}"#,
        ["1,2,3"]
    };

    iterable_single_where_finds_unique_match => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3, 4];
  print(nums.singleWhere((n) => n == 3));
}"#,
        ["3"]
    };

    iterable_single_where_on_singleton_list => {
        r#"void main() {
  Iterable<int> nums = [42];
  print(nums.singleWhere((n) => n > 0));
}"#,
        ["42"]
    };

    iterable_element_at_reads_middle_index => {
        r#"void main() {
  Iterable<int> nums = [10, 20, 30, 40];
  print(nums.elementAt(2));
}"#,
        ["30"]
    };

    iterable_element_at_reads_first_index => {
        r#"void main() {
  Iterable<String> words = ['a', 'b', 'c'];
  print(words.elementAt(0));
}"#,
        ["a"]
    };

    iterable_element_at_on_map_entries => {
        r#"void main() {
  var m = {1: 'one', 2: 'two'};
  var entries = m.entries;
  print(entries.elementAt(0).key);
}"#,
        ["1"]
    };

    iterable_fold_on_set_union => {
        r#"void main() {
  Iterable<int> s = {1, 2, 3};
  print(s.fold(0, (acc, n) => acc + n));
}"#,
        ["6"]
    };

    iterable_reduce_after_skip => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3, 4, 5];
  print(nums.skip(2).reduce((a, b) => a + b));
}"#,
        ["12"]
    };

    iterable_fold_after_where => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3, 4, 5, 6];
  print(nums.where((n) => n % 2 == 0).fold(0, (acc, n) => acc + n));
}"#,
        ["12"]
    };

    iterable_expand_then_take => {
        r#"void main() {
  Iterable<List<int>> nested = [[1, 2], [3, 4, 5]];
  print(nested.expand((p) => p).take(3).join(','));
}"#,
        ["1,2,3"]
    };

    iterable_followed_by_then_to_list => {
        r#"void main() {
  Iterable<int> a = [1];
  var list = a.followedBy([2, 3]).toList();
  print(list.length);
  print(list[2]);
}"#,
        ["3", "3"]
    };

    iterable_skip_while_then_take => {
        r#"void main() {
  Iterable<int> nums = [1, 1, 2, 2, 3, 3];
  print(nums.skipWhile((n) => n == 1).take(2).join(','));
}"#,
        ["2,2"]
    };

    iterable_take_while_then_reduce => {
        r#"void main() {
  Iterable<int> nums = [2, 4, 6, 1, 3];
  print(nums.takeWhile((n) => n % 2 == 0).reduce((a, b) => a + b));
}"#,
        ["12"]
    };

    iterable_cast_then_map => {
        r#"void main() {
  Iterable<Object> objs = [1, 2, 3];
  print(objs.cast<int>().map((n) => n + 1).join(','));
}"#,
        ["2,3,4"]
    };

    iterable_to_set_then_length => {
        r#"void main() {
  Iterable<int> nums = [1, 1, 2, 2, 3];
  print(nums.toSet().length);
}"#,
        ["3"]
    };

    iterable_to_list_growable_from_literal => {
        r#"void main() {
  Iterable<int> it = [1, 2];
  var list = it.toList();
  list.add(3);
  print(list.length);
}"#,
        ["3"]
    };

    iterable_single_where_after_map => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3];
  print(nums.map((n) => n * 10).singleWhere((n) => n == 20));
}"#,
        ["20"]
    };

    iterable_element_at_after_skip => {
        r#"void main() {
  Iterable<int> nums = [0, 1, 2, 3, 4];
  print(nums.skip(3).elementAt(0));
}"#,
        ["3"]
    };

    iterable_fold_on_string_split => {
        r#"void main() {
  Iterable<String> chars = 'abc'.split('');
  print(chars.fold(0, (acc, c) => acc + c.length));
}"#,
        ["3"]
    };

    iterable_reduce_on_followed_by_result => {
        r#"void main() {
  Iterable<int> a = [1, 2];
  print(a.followedBy([3]).reduce((x, y) => x + y));
}"#,
        ["6"]
    };

    iterable_expand_count_after_flatten => {
        r#"void main() {
  Iterable<List<int>> groups = [[1], [2, 3], [4, 5, 6]];
  print(groups.expand((g) => g).length);
}"#,
        ["6"]
    };

    iterable_skip_while_on_strings => {
        r#"void main() {
  Iterable<String> words = ['', '', 'hi', 'bye'];
  print(words.skipWhile((w) => w.isEmpty).join(','));
}"#,
        ["hi,bye"]
    };

    iterable_take_while_on_strings => {
        r#"void main() {
  Iterable<String> words = ['a', 'ab', 'abc', 'b'];
  print(words.takeWhile((w) => w.startsWith('a')).join('|'));
}"#,
        ["a|ab|abc"]
    };

    iterable_to_set_from_where_result => {
        r#"void main() {
  Iterable<int> nums = [1, 2, 3, 4, 5, 6];
  var s = nums.where((n) => n % 2 == 0).toSet();
  print(s.join(','));
}"#,
        ["2,4,6"]
    };

    iterable_element_at_last_position => {
        r#"void main() {
  Iterable<int> nums = [5, 6, 7];
  print(nums.elementAt(nums.length - 1));
}"#,
        ["7"]
    };

    iterable_fold_multiply_with_seed_one => {
        r#"void main() {
  Iterable<int> nums = [2, 3, 4];
  print(nums.fold(1, (acc, n) => acc * n));
}"#,
        ["24"]
    };

    iterable_map_values_to_list_length => {
        r#"void main() {
  var m = {'x': 10, 'y': 20, 'z': 30};
  print(m.values.toList().length);
  print(m.values.elementAt(1));
}"#,
        ["3", "20"]
    };

    iterable_chain_skip_take_to_list => {
        r#"void main() {
  Iterable<int> nums = [0, 1, 2, 3, 4, 5];
  var list = nums.skip(1).take(3).toList();
  print(list.join(','));
}"#,
        ["1,2,3"]
    };
}
