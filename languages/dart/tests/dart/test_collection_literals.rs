//! Collection-if and collection-for in list, map, and set literals.

dart_cases! {
    list_collection_if_true_adds_element => {
        r#"void main() {
  var show = true;
  var list = [1, 2, if (show) 3];
  print(list.length);
  print(list.last);
}"#,
        ["3", "3"]
    };

    list_collection_if_false_omits_element => {
        r#"void main() {
  var show = false;
  var list = [1, 2, if (show) 3];
  print(list.length);
  print(list.join(','));
}"#,
        ["2", "1,2"]
    };

    list_collection_if_else_true_branch => {
        r#"void main() {
  var flag = true;
  var list = [if (flag) 10 else 99];
  print(list[0]);
}"#,
        ["10"]
    };

    list_collection_if_else_false_branch => {
        r#"void main() {
  var flag = false;
  var list = [if (flag) 10 else 99];
  print(list[0]);
}"#,
        ["99"]
    };

    list_multiple_collection_if_in_one_literal => {
        r#"void main() {
  var a = true;
  var b = false;
  var list = [1, if (a) 2, if (b) 3, if (!b) 4];
  print(list.join(','));
}"#,
        ["1,2,4"]
    };

    list_collection_if_with_comparison_condition => {
        r#"void main() {
  var n = 7;
  var list = [1, if (n > 5) 2];
  print(list.length);
}"#,
        ["2"]
    };

    list_collection_if_nested_condition => {
        r#"void main() {
  var outer = true;
  var inner = false;
  var list = [if (outer) if (inner) 1 else 2];
  print(list[0]);
}"#,
        ["2"]
    };

    map_collection_if_true_adds_entry => {
        r#"void main() {
  var admin = true;
  var m = {'home': 1, if (admin) 'settings': 2};
  print(m.length);
  print(m['settings']);
}"#,
        ["2", "2"]
    };

    map_collection_if_false_skips_entry => {
        r#"void main() {
  var admin = false;
  var m = {'home': 1, if (admin) 'settings': 2};
  print(m.length);
  print(m.containsKey('settings'));
}"#,
        ["1", "false"]
    };

    map_collection_if_else_picks_value => {
        r#"void main() {
  var debug = false;
  var m = {'level': if (debug) 0 else 1};
  print(m['level']);
}"#,
        ["1"]
    };

    map_collection_if_multiple_conditional_keys => {
        r#"void main() {
  var a = true;
  var b = true;
  var m = {if (a) 'x': 1, if (b) 'y': 2};
  print(m.keys.join(','));
}"#,
        ["x,y"]
    };

    set_collection_if_true_adds_member => {
        r#"void main() {
  var extra = true;
  var s = {1, 2, if (extra) 3};
  print(s.length);
  print(s.contains(3));
}"#,
        ["3", "true"]
    };

    set_collection_if_false_omits_member => {
        r#"void main() {
  var extra = false;
  var s = {1, 2, if (extra) 3};
  print(s.length);
}"#,
        ["2"]
    };

    set_collection_if_else_selects_element => {
        r#"void main() {
  var pickA = false;
  var s = {if (pickA) 10 else 20};
  print(s.first);
}"#,
        ["20"]
    };

    list_collection_for_classic_loop => {
        r#"void main() {
  var list = [for (var i = 0; i < 4; i++) i];
  print(list.join(','));
}"#,
        ["0,1,2,3"]
    };

    list_collection_for_in_over_list => {
        r#"void main() {
  var src = [1, 2, 3];
  var doubled = [for (var x in src) x * 2];
  print(doubled.join('-'));
}"#,
        ["2-4-6"]
    };

    list_collection_for_generates_squares => {
        r#"void main() {
  var squares = [for (var i = 1; i <= 4; i++) i * i];
  print(squares.join(','));
}"#,
        ["1,4,9,16"]
    };

    list_collection_for_with_if_filter => {
        r#"void main() {
  var evens = [for (var i = 0; i < 6; i++) if (i.isEven) i];
  print(evens.join(','));
}"#,
        ["0,2,4"]
    };

    list_collection_for_in_strings_transform => {
        r#"void main() {
  var words = ['a', 'b'];
  var upper = [for (var w in words) w.toUpperCase()];
  print(upper.join(','));
}"#,
        ["A,B"]
    };

    list_collection_for_empty_range => {
        r#"void main() {
  var list = [for (var i = 0; i < 0; i++) i];
  print(list.length);
}"#,
        ["0"]
    };

    map_collection_for_index_to_square => {
        r#"void main() {
  var m = {for (var i = 0; i < 3; i++) i: i * i};
  print(m[2]);
  print(m.length);
}"#,
        ["4", "3"]
    };

    map_collection_for_in_over_list => {
        r#"void main() {
  var names = ['a', 'b'];
  var m = {for (var i = 0; i < names.length; i++) names[i]: i};
  print(m['a']);
  print(m['b']);
}"#,
        ["0", "1"]
    };

    map_collection_for_with_if_filter => {
        r#"void main() {
  var m = {for (var i = 0; i < 5; i++) if (i.isOdd) i: i * 10};
  print(m[1]);
  print(m.length);
}"#,
        ["10", "2"]
    };

    set_collection_for_classic_loop => {
        r#"void main() {
  var s = {for (var i = 0; i < 3; i++) i};
  print(s.length);
}"#,
        ["3"]
    };

    set_collection_for_in_deduplicates => {
        r#"void main() {
  var src = [1, 1, 2, 2];
  var s = {for (var x in src) x};
  print(s.length);
}"#,
        ["2"]
    };

    set_collection_for_with_if_filter => {
        r#"void main() {
  var s = {for (var i = 0; i < 6; i++) if (i % 2 == 0) i};
  print(s.join(','));
}"#,
        ["0,2,4"]
    };

    list_collection_for_with_step_increment => {
        r#"void main() {
  var list = [for (var i = 0; i < 10; i += 3) i];
  print(list.join(','));
}"#,
        ["0,3,6,9"]
    };

    list_collection_if_and_for_combined => {
        r#"void main() {
  var includeExtra = true;
  var list = [for (var i = 1; i <= 3; i++) i, if (includeExtra) 99];
  print(list.join(','));
}"#,
        ["1,2,3,99"]
    };

    map_collection_for_string_keys_from_numbers => {
        r#"void main() {
  var m = {for (var i = 1; i <= 2; i++) 'k$i': i};
  print(m['k1']);
  print(m['k2']);
}"#,
        ["1", "2"]
    };

    list_collection_for_nested_in_expression => {
        r#"void main() {
  var rows = [for (var r = 0; r < 2; r++) r];
  var flat = [for (var r in rows) for (var c = 0; c < 2; c++) r * 10 + c];
  print(flat.join(','));
}"#,
        ["0,1,10,11"]
    };

    list_collection_if_variable_condition_changes => {
        r#"void main() {
  var flag = false;
  var a = [if (flag) 1];
  flag = true;
  var b = [if (flag) 2];
  print(a.length);
  print(b[0]);
}"#,
        ["0", "2"]
    };

    map_collection_if_else_two_different_keys => {
        r#"void main() {
  var prod = true;
  var m = {if (prod) 'env': 'prod' else 'env': 'dev'};
  print(m['env']);
}"#,
        ["prod"]
    };

    set_collection_for_from_range_with_offset => {
        r#"void main() {
  var s = {for (var i = 0; i < 3; i++) i + 10};
  print(s.contains(12));
  print(s.length);
}"#,
        ["true", "3"]
    };

    list_collection_for_in_characters => {
        r#"void main() {
  var chars = [for (var c in 'hi') c];
  print(chars.join(''));
}"#,
        ["hi"]
    };

    map_collection_for_empty_range => {
        r#"void main() {
  var m = {for (var i = 0; i < 0; i++) i: i};
  print(m.isEmpty);
}"#,
        ["true"]
    };

    list_collection_if_false_leaves_only_fixed_elements => {
        r#"void main() {
  var items = ['a', 'b', if (false) 'c', 'd'];
  print(items.join(','));
}"#,
        ["a,b,d"]
    };

    set_collection_if_true_with_existing_members => {
        r#"void main() {
  var add = true;
  var s = {1, 2, if (add) 3, if (add) 2};
  print(s.length);
}"#,
        ["3"]
    };

    list_collection_for_builds_pairs_via_expression => {
        r#"void main() {
  var pairs = [for (var i = 0; i < 3; i++) '$i:${i * 2}'];
  print(pairs[1]);
}"#,
        ["1:2"]
    };

    map_collection_for_in_with_string_values => {
        r#"void main() {
  var keys = ['x', 'y'];
  var m = {for (var k in keys) k: k.length};
  print(m['x']);
  print(m['y']);
}"#,
        ["1", "1"]
    };

    list_collection_for_descending_manual => {
        r#"void main() {
  var list = [for (var i = 3; i >= 1; i--) i];
  print(list.join(','));
}"#,
        ["3,2,1"]
    };
}
