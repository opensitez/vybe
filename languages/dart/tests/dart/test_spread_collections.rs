//! Spread operators in list, map, and set literals, including null-aware `...?`.

dart_cases! {
    list_spread_appends_after_source_elements => {
        r#"void main() {
  var a = [1, 2];
  var b = [...a, 3, 4];
  print(b.join(','));
}"#,
        ["1,2,3,4"]
    };

    list_spread_prepends_before_literals => {
        r#"void main() {
  var tail = [3, 4];
  var all = [1, 2, ...tail];
  print(all.first);
  print(all.last);
}"#,
        ["1", "4"]
    };

    list_spread_inserts_in_middle => {
        r#"void main() {
  var left = [1, 2];
  var right = [5, 6];
  var mid = [...left, 3, 4, ...right];
  print(mid.join('-'));
}"#,
        ["1-2-3-4-5-6"]
    };

    list_spread_from_empty_list_adds_nothing => {
        r#"void main() {
  var empty = <int>[];
  var merged = [...empty, 1, 2];
  print(merged.length);
  print(merged.join(','));
}"#,
        ["2", "1,2"]
    };

    list_spread_into_empty_result => {
        r#"void main() {
  var src = [7, 8];
  var dst = <int>[...src];
  print(dst.join(','));
}"#,
        ["7,8"]
    };

    list_spread_combines_three_sources => {
        r#"void main() {
  var a = [1];
  var b = [2];
  var c = [3];
  var all = [...a, ...b, ...c];
  print(all.join(''));
}"#,
        ["123"]
    };

    list_spread_preserves_element_order => {
        r#"void main() {
  var first = [10, 20];
  var second = [30, 40];
  var merged = [...first, ...second];
  print(merged[0]);
  print(merged[3]);
}"#,
        ["10", "40"]
    };

    list_spread_with_typed_source => {
        r#"void main() {
  List<int> nums = [4, 5];
  var out = <int>[0, ...nums];
  print(out.join(','));
}"#,
        ["0,4,5"]
    };

    list_spread_single_element_source => {
        r#"void main() {
  var one = [99];
  var out = [...one, 100];
  print(out.length);
  print(out[0]);
  print(out[1]);
}"#,
        ["2", "99", "100"]
    };

    list_spread_nested_variable_rebuild => {
        r#"void main() {
  var inner = [2, 3];
  var outer = [1, ...inner, 4];
  print(outer.join(','));
}"#,
        ["1,2,3,4"]
    };

    map_spread_merges_second_into_first => {
        r#"void main() {
  var a = {'x': 1};
  var b = {'y': 2, ...a};
  print(b['x']);
  print(b['y']);
  print(b.length);
}"#,
        ["1", "2", "2"]
    };

    map_spread_combines_two_maps => {
        r#"void main() {
  var left = {'a': 1};
  var right = {'b': 2};
  var merged = {...left, ...right};
  print(merged.keys.join(','));
}"#,
        ["a,b"]
    };

    map_spread_later_key_overwrites_earlier => {
        r#"void main() {
  var first = {'k': 1};
  var second = {'k': 9};
  var merged = {...first, ...second};
  print(merged['k']);
}"#,
        ["9"]
    };

    map_spread_from_empty_map => {
        r#"void main() {
  var empty = <String, int>{};
  var out = {...empty, 'a': 1};
  print(out.length);
  print(out['a']);
}"#,
        ["1", "1"]
    };

    map_spread_with_literal_entries_mixed => {
        r#"void main() {
  var base = {'a': 1};
  var out = {'z': 0, ...base, 'b': 2};
  print(out['z']);
  print(out['a']);
  print(out['b']);
}"#,
        ["0", "1", "2"]
    };

    map_spread_typed_empty_into_populated => {
        r#"void main() {
  Map<String, int> extra = {'m': 3};
  var out = <String, int>{'n': 4, ...extra};
  print(out.length);
  print(out['m']);
}"#,
        ["2", "3"]
    };

    map_spread_three_sources => {
        r#"void main() {
  var a = {'a': 1};
  var b = {'b': 2};
  var c = {'c': 3};
  var all = {...a, ...b, ...c};
  print(all.length);
}"#,
        ["3"]
    };

    set_spread_merges_two_sets => {
        r#"void main() {
  var a = {1, 2};
  var b = {3, 4};
  var merged = {...a, ...b};
  print(merged.length);
  print(merged.contains(3));
}"#,
        ["4", "true"]
    };

    set_spread_deduplicates_overlapping_elements => {
        r#"void main() {
  var a = {1, 2, 3};
  var b = {3, 4};
  var merged = {...a, ...b};
  print(merged.length);
}"#,
        ["4"]
    };

    set_spread_from_empty_set => {
        r#"void main() {
  var empty = <int>{};
  var out = {...empty, 1, 2};
  print(out.length);
}"#,
        ["2"]
    };

    set_spread_with_literal_values_mixed => {
        r#"void main() {
  var base = {10, 20};
  var out = {0, ...base, 30};
  print(out.length);
}"#,
        ["4"]
    };

    set_spread_typed_source => {
        r#"void main() {
  Set<int> src = {5, 6};
  var out = <int>{1, ...src};
  print(out.contains(6));
  print(out.length);
}"#,
        ["true", "3"]
    };

    set_spread_preserves_uniqueness_from_source => {
        r#"void main() {
  var src = {1, 1, 2};
  var out = {...src, 3};
  print(out.length);
}"#,
        ["3"]
    };

    null_aware_list_spread_skips_null_source => {
        r#"void main() {
  List<int>? missing = null;
  var out = [...?missing, 1, 2];
  print(out.join(','));
}"#,
        ["1,2"]
    };

    null_aware_list_spread_includes_non_null_source => {
        r#"void main() {
  List<int>? present = [3, 4];
  var out = [1, 2, ...?present];
  print(out.join(','));
}"#,
        ["1,2,3,4"]
    };

    null_aware_map_spread_skips_null_source => {
        r#"void main() {
  Map<String, int>? missing = null;
  var out = {...?missing, 'a': 1};
  print(out.length);
  print(out['a']);
}"#,
        ["1", "1"]
    };

    null_aware_map_spread_includes_non_null_source => {
        r#"void main() {
  Map<String, int>? extra = {'b': 2};
  var out = {'a': 1, ...?extra};
  print(out.length);
  print(out['b']);
}"#,
        ["2", "2"]
    };

    null_aware_set_spread_skips_null_source => {
        r#"void main() {
  Set<int>? missing = null;
  var out = {...?missing, 1, 2};
  print(out.length);
}"#,
        ["2"]
    };

    null_aware_set_spread_includes_non_null_source => {
        r#"void main() {
  Set<int>? extra = {3, 4};
  var out = {1, 2, ...?extra};
  print(out.length);
}"#,
        ["4"]
    };

    null_aware_spread_mixed_null_and_present_lists => {
        r#"void main() {
  List<int>? a = null;
  List<int>? b = [2, 3];
  var out = [1, ...?a, ...?b, 4];
  print(out.join(','));
}"#,
        ["1,2,3,4"]
    };

    null_aware_spread_both_null_adds_only_literals => {
        r#"void main() {
  List<int>? a = null;
  List<int>? b = null;
  var out = [...?a, ...?b, 9];
  print(out.length);
  print(out[0]);
}"#,
        ["1", "9"]
    };

    null_aware_list_spread_on_reassigned_nullable => {
        r#"void main() {
  List<int>? data = null;
  var before = [...?data, 1];
  data = [2];
  var after = [...?data, 3];
  print(before.join(','));
  print(after.join(','));
}"#,
        ["1", "2,3"]
    };

    list_spread_then_index_access => {
        r#"void main() {
  var src = [5, 6];
  var dst = [1, ...src, 7];
  print(dst[1]);
  print(dst[2]);
}"#,
        ["5", "6"]
    };

    map_spread_key_collision_three_layers => {
        r#"void main() {
  var a = {'k': 1};
  var b = {'k': 2};
  var c = {'k': 3};
  var out = {...a, ...b, ...c};
  print(out['k']);
}"#,
        ["3"]
    };

    set_spread_with_duplicate_literals_and_spread => {
        r#"void main() {
  var src = {2, 3};
  var out = {1, 2, ...src, 3};
  print(out.length);
}"#,
        ["3"]
    };

    list_spread_chained_length_check => {
        r#"void main() {
  var a = [1, 2];
  var b = [3];
  var c = [4, 5, 6];
  var all = [...a, ...b, ...c];
  print(all.length);
}"#,
        ["6"]
    };

    map_spread_empty_into_empty_then_add => {
        r#"void main() {
  var e1 = <String, int>{};
  var e2 = <String, int>{};
  var out = {...e1, ...e2, 'x': 1};
  print(out.length);
}"#,
        ["1"]
    };

    null_aware_map_spread_null_between_literals => {
        r#"void main() {
  Map<String, int>? mid = null;
  var out = {'a': 1, ...?mid, 'b': 2};
  print(out.keys.join(','));
}"#,
        ["a,b"]
    };

    set_spread_single_element_source => {
        r#"void main() {
  var src = {42};
  var out = {...src, 43};
  print(out.length);
  print(out.contains(42));
}"#,
        ["2", "true"]
    };

    list_null_aware_spread_only_source_null => {
        r#"void main() {
  List<int>? src = null;
  var out = <int>[...?src];
  print(out.isEmpty);
}"#,
        ["true"]
    };
}
