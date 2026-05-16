use super::helpers::{compile_ok, run_prints};

// ── Spread operator ─────────────────────────────────────────

#[test] fn spread_list() { compile_ok("var a = [1, 2]; var b = [...a, 3, 4];"); }
#[test] fn spread_combine() { compile_ok("var a = [1, 2]; var b = [3, 4]; var c = [...a, ...b];"); }
#[test] fn spread_null_aware() { compile_ok("List? a = null; var b = [...?a, 1, 2];"); }
#[test] fn spread_in_call() { compile_ok("void f(int x, int y) {} var args = [1, 2];"); }

#[test] fn spread_result() {
    let out = run_prints("void main() { var a = [1, 2]; var b = [...a, 3]; print(b.length); }");
    assert_eq!(out, ["3"]);
}

#[test] fn spread_map() { compile_ok("var a = {'x': 1}; var b = {'y': 2, ...a};"); }

// ── Collection if ────────────────────────────────────────────

#[test] fn collection_if() { compile_ok("var showExtra = true; var list = [1, 2, if (showExtra) 3];"); }
#[test] fn collection_if_else() { compile_ok("var flag = false; var list = [1, if (flag) 2 else 99];"); }
#[test] fn collection_if_string() { compile_ok("var admin = true; var items = ['home', if (admin) 'settings'];"); }

#[test] fn collection_if_result() {
    let out = run_prints("void main() { var show = true; var list = [1, 2, if (show) 3]; print(list.length); }");
    assert_eq!(out, ["3"]);
}

#[test] fn collection_if_false_result() {
    let out = run_prints("void main() { var show = false; var list = [1, 2, if (show) 3]; print(list.length); }");
    assert_eq!(out, ["2"]);
}

// ── Collection for ───────────────────────────────────────────

#[test] fn collection_for() { compile_ok("var list = [for (var i = 0; i < 5; i++) i];"); }
#[test] fn collection_for_expr() { compile_ok("var squares = [for (var i = 1; i <= 5; i++) i * i];"); }
#[test] fn collection_for_in() { compile_ok("var source = [1, 2, 3]; var doubled = [for (var x in source) x * 2];"); }
#[test] fn collection_for_strings() { compile_ok("var words = ['a', 'b', 'c']; var upper = [for (var w in words) w.toUpperCase()];"); }

#[test] fn collection_for_result() {
    let out = run_prints("void main() { var list = [for (var i = 0; i < 3; i++) i]; print(list.length); }");
    assert_eq!(out, ["3"]);
}

#[test] fn collection_for_map() { compile_ok("var m = {for (var i = 0; i < 3; i++) i: i * i};"); }

// ── Set type ────────────────────────────────────────────────

#[test] fn set_literal() { compile_ok("var s = {1, 2, 3};"); }
#[test] fn set_from_list() { compile_ok("var s = Set.from([1, 2, 2, 3]);"); }
#[test] fn set_typed() { compile_ok("Set<int> s = {1, 2, 3};"); }
#[test] fn set_add() { compile_ok("var s = <int>{}; s.add(1); s.add(2);"); }
#[test] fn set_contains() { compile_ok("var s = {1, 2, 3}; var b = s.contains(2);"); }
#[test] fn set_remove() { compile_ok("var s = {1, 2, 3}; s.remove(2);"); }
#[test] fn set_length() { compile_ok("var s = {1, 2, 3}; var n = s.length;"); }
#[test] fn set_is_empty() { compile_ok("var s = <int>{}; var b = s.isEmpty;"); }
#[test] fn set_union() { compile_ok("var a = {1, 2}; var b = {2, 3}; var c = a.union(b);"); }
#[test] fn set_intersection() { compile_ok("var a = {1, 2, 3}; var b = {2, 3, 4}; var c = a.intersection(b);"); }
#[test] fn set_difference() { compile_ok("var a = {1, 2, 3}; var b = {2}; var c = a.difference(b);"); }
#[test] fn set_foreach() { compile_ok("var s = {1, 2, 3}; s.forEach((e) => print(e));"); }
#[test] fn set_to_list() { compile_ok("var s = {1, 2, 3}; var list = s.toList();"); }

// ── List advanced operations ─────────────────────────────────

#[test] fn list_insert() { compile_ok("var list = [1, 2, 3]; list.insert(1, 99);"); }
#[test] fn list_remove_at() { compile_ok("var list = [1, 2, 3]; list.removeAt(0);"); }
#[test] fn list_remove() { compile_ok("var list = [1, 2, 3]; list.remove(2);"); }
#[test] fn list_remove_where() { compile_ok("var list = [1, 2, 3, 4]; list.removeWhere((e) => e % 2 == 0);"); }
#[test] fn list_sort() { compile_ok("var list = [3, 1, 2]; list.sort();"); }
#[test] fn list_sort_custom() { compile_ok("var list = [3, 1, 2]; list.sort((a, b) => a.compareTo(b));"); }
#[test] fn list_index_of() { compile_ok("var list = [10, 20, 30]; var i = list.indexOf(20);"); }
#[test] fn list_contains() { compile_ok("var list = [1, 2, 3]; var b = list.contains(2);"); }
#[test] fn list_clear() { compile_ok("var list = [1, 2, 3]; list.clear();"); }
#[test] fn list_sublist() { compile_ok("var list = [1, 2, 3, 4, 5]; var sub = list.sublist(1, 3);"); }
#[test] fn list_skip() { compile_ok("var list = [1, 2, 3, 4]; var rest = list.skip(2).toList();"); }
#[test] fn list_take() { compile_ok("var list = [1, 2, 3, 4]; var first2 = list.take(2).toList();"); }
#[test] fn list_fold() { compile_ok("var sum = [1, 2, 3, 4].fold(0, (acc, e) => acc + e);"); }
#[test] fn list_expand() { compile_ok("var flat = [[1,2],[3,4]].expand((e) => e).toList();"); }
#[test] fn list_generate() { compile_ok("var list = List.generate(5, (i) => i * 2);"); }
#[test] fn list_filled() { compile_ok("var list = List.filled(3, 0);"); }
#[test] fn list_add_all() { compile_ok("var a = [1, 2]; a.addAll([3, 4]);"); }
#[test] fn list_to_set() { compile_ok("var set = [1, 2, 2, 3].toSet();"); }

#[test] fn list_fold_result() {
    let out = run_prints("void main() { var sum = [1, 2, 3, 4].fold(0, (acc, e) => acc + e); print(sum); }");
    assert_eq!(out, ["10"]);
}

#[test] fn list_sort_result() {
    let out = run_prints("void main() { var list = [3, 1, 2]; list.sort(); print(list.first); }");
    assert_eq!(out, ["1"]);
}

// ── Map advanced operations ──────────────────────────────────

#[test] fn map_put_if_absent() { compile_ok("var m = {'a': 1}; m.putIfAbsent('b', () => 2);"); }
#[test] fn map_update() { compile_ok("var m = {'a': 1}; m.update('a', (v) => v + 1);"); }
#[test] fn map_remove() { compile_ok("var m = {'a': 1, 'b': 2}; m.remove('a');"); }
#[test] fn map_contains_key() { compile_ok("var m = {'a': 1}; var b = m.containsKey('a');"); }
#[test] fn map_contains_value() { compile_ok("var m = {'a': 1}; var b = m.containsValue(1);"); }
#[test] fn map_foreach() { compile_ok("var m = {'a': 1, 'b': 2}; m.forEach((k, v) => print('$k:$v'));"); }
#[test] fn map_keys() { compile_ok("var m = {'a': 1, 'b': 2}; var keys = m.keys.toList();"); }
#[test] fn map_values() { compile_ok("var m = {'a': 1, 'b': 2}; var vals = m.values.toList();"); }
#[test] fn map_entries() { compile_ok("var m = {'a': 1}; var entries = m.entries.toList();"); }
#[test] fn map_length() { compile_ok("var m = {'a': 1, 'b': 2}; var n = m.length;"); }
#[test] fn map_is_empty() { compile_ok("var m = <String, int>{}; var b = m.isEmpty;"); }
#[test] fn map_set_new_key() { compile_ok("var m = <String, int>{}; m['x'] = 42;"); }
#[test] fn map_typed() { compile_ok("Map<String, int> scores = {'Alice': 95, 'Bob': 87};"); }

#[test] fn map_contains_key_result() {
    let out = run_prints("void main() { var m = {'x': 1}; print(m.containsKey('x')); }");
    assert_eq!(out, ["true"]);
}

#[test] fn map_length_result() {
    let out = run_prints("void main() { var m = {'a': 1, 'b': 2}; print(m.length); }");
    assert_eq!(out, ["2"]);
}

// ── Iterable ─────────────────────────────────────────────────

#[test] fn iterable_to_list() { compile_ok("Iterable<int> it = [1, 2, 3]; var list = it.toList();"); }
#[test] fn iterable_to_set() { compile_ok("Iterable<int> it = [1, 2, 2]; var s = it.toSet();"); }
#[test] fn iterable_map() { compile_ok("Iterable<int> it = [1, 2, 3]; var r = it.map((e) => e * 2);"); }
#[test] fn iterable_where() { compile_ok("Iterable<int> it = [1, 2, 3]; var r = it.where((e) => e > 1);"); }
