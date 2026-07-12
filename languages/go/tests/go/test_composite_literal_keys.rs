//! Keyed composite literals: struct field keys, array/slice index keys,
//! map entry keys, inferred-length `[...]` arrays, and nested combinations.

go_run_cases! {
    // --- keyed struct literals ---
    struct_three_fields_keyed_reverse_order => (
        "package main; import \"fmt\"; type record struct { a int; b int; c int }; func main() { r := record{c: 3, b: 2, a: 1}; fmt.Println(r.a); fmt.Println(r.c); }",
        vec!["1", "3"]
    ),
    struct_partial_keyed_leaves_zero_values => (
        "package main; import \"fmt\"; type config struct { host string; port int; debug bool }; func main() { c := config{port: 8080}; fmt.Println(c.port); fmt.Println(c.debug); }",
        vec!["8080", "false"]
    ),
    struct_mixed_type_fields_all_keyed => (
        "package main; import \"fmt\"; type item struct { label string; count int; active bool }; func main() { it := item{active: true, label: \"vybe\", count: 7}; fmt.Println(it.label); fmt.Println(it.count); fmt.Println(it.active); }",
        vec!["vybe", "7", "true"]
    ),
    struct_keyed_embedded_inner_by_name => (
        "package main; import \"fmt\"; type inner struct { value int }; type outer struct { inner }; func main() { o := outer{inner: inner{value: 42}}; fmt.Println(o.value); }",
        vec!["42"]
    ),
    struct_nested_keyed_inner_fields_shuffled => (
        "package main; import \"fmt\"; type coord struct { x int; y int }; type rect struct { origin coord; size coord }; func main() { r := rect{size: coord{y: 4, x: 3}, origin: coord{x: 1, y: 2}}; fmt.Println(r.origin.x); fmt.Println(r.size.y); }",
        vec!["1", "4"]
    ),
    anonymous_struct_literal_keyed_fields => (
        "package main; import \"fmt\"; func main() { v := struct { id int; name string }{name: \"go\", id: 9}; fmt.Println(v.id); fmt.Println(v.name); }",
        vec!["9", "go"]
    ),
    struct_keyed_pointer_field_inline => (
        "package main; import \"fmt\"; type node struct { value int }; type holder struct { head *node }; func main() { h := holder{head: &node{value: 17}}; fmt.Println(h.head.value); }",
        vec!["17"]
    ),

    // --- keyed array literals ---
    array_keyed_index_zero_explicit => (
        "package main; import \"fmt\"; func main() { a := [5]int{0: 11, 4: 44}; fmt.Println(a[0]); fmt.Println(a[1]); fmt.Println(a[4]); }",
        vec!["11", "0", "44"]
    ),
    array_keyed_mixed_with_positional_continuation => (
        "package main; import \"fmt\"; func main() { a := [6]int{1: 10, 3: 30, 5}; fmt.Println(a[1]); fmt.Println(a[3]); fmt.Println(a[4]); fmt.Println(len(a)); }",
        vec!["10", "30", "5", "6"]
    ),
    array_inferred_length_with_keyed_indices => (
        "package main; import \"fmt\"; func main() { a := [...]int{3: 9, 4: 10}; fmt.Println(len(a)); fmt.Println(a[3]); fmt.Println(a[0]); }",
        vec!["5", "9", "0"]
    ),
    array_inferred_length_nested_2d => (
        "package main; import \"fmt\"; func main() { grid := [...][2]int{{0: 1, 1: 2}, {1: 4, 0: 3}}; fmt.Println(grid[1][0]); fmt.Println(grid[0][1]); }",
        vec!["3", "2"]
    ),
    array_of_structs_mixed_keyed_unkeyed_elements => (
        "package main; import \"fmt\"; type pair struct { left int; right int }; func main() { a := [3]pair{{right: 2, left: 1}, {3, 4}, pair{left: 5, right: 6}}; fmt.Println(a[0].left); fmt.Println(a[1].right); fmt.Println(a[2].left); }",
        vec!["1", "4", "5"]
    ),
    array_keyed_inside_keyed_struct_field => (
        "package main; import \"fmt\"; type box struct { data [4]int }; func main() { b := box{data: [4]int{0: 7, 3: 9}}; fmt.Println(b.data[0]); fmt.Println(b.data[2]); fmt.Println(b.data[3]); }",
        vec!["7", "0", "9"]
    ),

    // --- keyed slice literals & `[...]` ---
    slice_keyed_sparse_high_index => (
        "package main; import \"fmt\"; func main() { s := []int{10: 1}; fmt.Println(len(s)); fmt.Println(s[10]); fmt.Println(s[0]); }",
        vec!["11", "1", "0"]
    ),
    slice_keyed_mixed_with_positional => (
        "package main; import \"fmt\"; func main() { s := []int{1: 10, 20, 30}; fmt.Println(s[1]); fmt.Println(s[2]); fmt.Println(s[3]); fmt.Println(len(s)); }",
        vec!["10", "20", "30", "4"]
    ),
    slice_of_arrays_with_keyed_inner_indices => (
        "package main; import \"fmt\"; func main() { s := [][2]int{{0: 5, 1: 6}, {1: 8, 0: 7}}; fmt.Println(s[0][0]); fmt.Println(s[1][1]); }",
        vec!["5", "8"]
    ),
    inferred_array_string_elements_ellipsis => (
        "package main; import \"fmt\"; func main() { words := [...]string{\"go\", \"vybe\", \"keys\"}; fmt.Println(len(words)); fmt.Println(words[2]); }",
        vec!["3", "keys"]
    ),

    // --- keyed map literals ---
    map_negative_int_keys => (
        "package main; import \"fmt\"; func main() { m := map[int]string{-1: \"minus\", 0: \"zero\", 1: \"plus\"}; fmt.Println(m[-1]); fmt.Println(m[0]); }",
        vec!["minus", "zero"]
    ),
    map_rune_keys => (
        "package main; import \"fmt\"; func main() { m := map[rune]string{'a': \"alpha\", 'z': \"omega\"}; fmt.Println(m['a']); fmt.Println(m['z']); }",
        vec!["alpha", "omega"]
    ),
    map_value_struct_partial_keyed_fields => (
        "package main; import \"fmt\"; type point struct { x int; y int; label string }; func main() { m := map[string]point{\"p\": {y: 9, label: \"home\"}}; fmt.Println(m[\"p\"].y); fmt.Println(m[\"p\"].x); fmt.Println(m[\"p\"].label); }",
        vec!["9", "0", "home"]
    ),
    map_value_slice_with_keyed_indices => (
        "package main; import \"fmt\"; func main() { m := map[string][]int{\"data\": {0: 5, 2: 7}}; fmt.Println(len(m[\"data\"])); fmt.Println(m[\"data\"][0]); fmt.Println(m[\"data\"][2]); }",
        vec!["3", "5", "7"]
    ),
    map_nested_struct_and_slice_keys => (
        "package main; import \"fmt\"; type cell struct { n int }; func main() { m := map[string]struct { rows []cell }{ \"t\": {rows: []cell{{n: 1}, {n: 2}}} }; fmt.Println(m[\"t\"].rows[1].n); }",
        vec!["2"]
    ),

    // --- nested composite literals with keys ---
    nested_map_array_struct_keyed_chain => (
        "package main; import \"fmt\"; type item struct { id int }; func main() { data := map[string][]item{\"batch\": {{id: 10}, {id: 20}}}; fmt.Println(data[\"batch\"][0].id); fmt.Println(data[\"batch\"][1].id); }",
        vec!["10", "20"]
    ),
    nested_struct_map_array_all_keyed => (
        "package main; import \"fmt\"; type entry struct { scores []int }; type table struct { rows map[string]entry }; func main() { t := table{rows: map[string]entry{\"a\": {scores: []int{0: 100, 2: 300}}}}; fmt.Println(t.rows[\"a\"].scores[0]); fmt.Println(t.rows[\"a\"].scores[2]); }",
        vec!["100", "300"]
    ),
    nested_four_level_keyed_composite => (
        "package main; import \"fmt\"; type leaf struct { v int }; type branch struct { leaves []leaf }; type tree struct { parts []branch }; func main() { tr := tree{parts: []branch{{leaves: []leaf{{v: 99}}}}}; fmt.Println(tr.parts[0].leaves[0].v); }",
        vec!["99"]
    ),
    slice_of_maps_with_keyed_struct_values => (
        "package main; import \"fmt\"; type pair struct { a int; b int }; func main() { s := []map[string]pair{{\"x\": {b: 2, a: 1}}, {\"y\": pair{a: 3, b: 4}}}; fmt.Println(s[0][\"x\"].a); fmt.Println(s[1][\"y\"].b); }",
        vec!["1", "4"]
    ),
    pointer_to_keyed_struct_literal => (
        "package main; import \"fmt\"; type metric struct { name string; value int }; func main() { m := &metric{value: 42, name: \"latency\"}; fmt.Println(m.name); fmt.Println(m.value); }",
        vec!["latency", "42"]
    ),
}

go_compile_cases! {
    array_of_maps_keyed_inner_entries => "package main; func main() { _ = [2]map[string]int{{\"a\": 1}, {\"b\": 2}} }",
    map_of_arrays_keyed_slice_values => "package main; func main() { _ = map[string][3]int{\"row\": {0: 1, 2: 3}} }",
    struct_with_keyed_map_and_array_fields => "package main; type bundle struct { tags map[string]int; ids [2]int }; func main() { _ = bundle{tags: map[string]int{\"x\": 1}, ids: [2]int{1: 9}} }",
    nested_anonymous_struct_all_keyed_compile => "package main; func main() { _ = struct { outer struct { n int } }{outer: struct { n int }{n: 5}} }",
    slice_keyed_struct_elements_compile => "package main; type node struct { id int }; func main() { _ = []node{{id: 1}, {id: 2}} }",
}
