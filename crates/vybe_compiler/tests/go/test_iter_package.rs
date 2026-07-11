//! Go 1.23 `iter` package and iterator helpers: Seq, Pull, Pull2, slices.All,
//! slices.Values over map values via maps.Values — compile smoke plus runtime
//! where the VM can execute iterator loops synchronously.

go_run_cases! {
    iter_slices_all_index_sum => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{10, 20, 30}; sum := 0; for i := range slices.All(s) { sum += s[i] }; fmt.Println(sum) }",
        vec!["60"]
    ),
    iter_slices_all_empty_slice => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{}; count := 0; for range slices.All(s) { count++ }; fmt.Println(count) }",
        vec!["0"]
    ),
    iter_slices_all_single_element => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{42}; idx := -1; for i := range slices.All(s) { idx = i }; fmt.Println(idx) }",
        vec!["0"]
    ),
    iter_slices_all_collect_indices => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []string{\"a\", \"b\", \"c\"}; n := 0; for i := range slices.All(s) { if s[i] != \"\" { n++ } }; fmt.Println(n) }",
        vec!["3"]
    ),
    iter_maps_values_sum => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[string]int{\"a\": 1, \"b\": 2, \"c\": 3}; sum := 0; for v := range maps.Values(m) { sum += v }; fmt.Println(sum) }",
        vec!["6"]
    ),
    iter_maps_values_empty_map => (
        "package main; import \"fmt\"; import \"maps\"; func main() { count := 0; for range maps.Values(map[int]int{}) { count++ }; fmt.Println(count) }",
        vec!["0"]
    ),
    iter_maps_values_count_strings => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]string{1: \"go\", 2: \"vybe\"}; n := 0; for v := range maps.Values(m) { n += len(v) }; fmt.Println(n) }",
        vec!["6"]
    ),
    iter_pull_manual_next_stop => (
        "package main; import \"fmt\"; import \"iter\"; func main() { seq := func(yield func(int) bool) { yield(1); yield(2); yield(3) }; next, stop := iter.Pull(seq); defer stop(); v1, ok1 := next(); v2, ok2 := next(); _, ok3 := next(); _, ok4 := next(); fmt.Println(v1); fmt.Println(v2); fmt.Println(ok1 && ok2 && ok3 && !ok4) }",
        vec!["1", "2", "true"]
    ),
    iter_pull_early_stop => (
        "package main; import \"fmt\"; import \"iter\"; func main() { seq := func(yield func(int) bool) { if !yield(10) { return }; yield(20) }; next, stop := iter.Pull(seq); defer stop(); v, ok := next(); stop(); fmt.Println(v); fmt.Println(ok) }",
        vec!["10", "true"]
    ),
    iter_pull2_key_value_pairs => (
        "package main; import \"fmt\"; import \"iter\"; func main() { seq := func(yield func(int, string) bool) { yield(1, \"a\"); yield(2, \"b\") }; next, stop := iter.Pull2(seq); defer stop(); k, v, ok := next(); fmt.Println(k); fmt.Println(v); fmt.Println(ok) }",
        vec!["1", "a", "true"]
    ),
    iter_pull2_exhausted => (
        "package main; import \"fmt\"; import \"iter\"; func main() { seq := func(yield func(int, int) bool) { yield(0, 0) }; next, stop := iter.Pull2(seq); defer stop(); _, _, ok1 := next(); _, _, ok2 := next(); fmt.Println(ok1); fmt.Println(ok2) }",
        vec!["true", "false"]
    ),
    iter_seq_range_over_custom => (
        "package main; import \"fmt\"; func main() { seq := func(yield func(int) bool) { for i := 1; i <= 3; i++ { if !yield(i) { return } } }; sum := 0; for v := range seq { sum += v }; fmt.Println(sum) }",
        vec!["6"]
    ),
    iter_seq_break_stops_yield => (
        "package main; import \"fmt\"; func main() { count := 0; seq := func(yield func(int) bool) { for i := 0; i < 100; i++ { if !yield(i) { return }; count++ } }; for v := range seq { if v == 2 { break } }; fmt.Println(count) }",
        vec!["3"]
    ),
    iter_slices_values_over_map => (
        "package main; import \"fmt\"; import \"slices\"; func main() { m := map[int]int{1: 10, 2: 20}; sum := 0; for v := range slices.Values(m) { sum += v }; fmt.Println(sum) }",
        vec!["30"]
    ),
    iter_slices_values_empty => (
        "package main; import \"fmt\"; import \"slices\"; func main() { n := 0; for range slices.Values(map[string]int{}) { n++ }; fmt.Println(n) }",
        vec!["0"]
    ),
    iter_slices_all_string_slice => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []string{\"x\", \"y\"}; acc := \"\"; for i := range slices.All(s) { acc += s[i] }; fmt.Println(acc) }",
        vec!["xy"]
    ),
    iter_pull_three_values => (
        "package main; import \"fmt\"; import \"iter\"; func main() { seq := func(yield func(int) bool) { yield(5); yield(6); yield(7) }; next, stop := iter.Pull(seq); defer stop(); a, _ := next(); b, _ := next(); c, _ := next(); fmt.Println(a + b + c) }",
        vec!["18"]
    ),
    iter_maps_values_bool_map => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[string]bool{\"a\": true, \"b\": false, \"c\": true}; trues := 0; for v := range maps.Values(m) { if v { trues++ } }; fmt.Println(trues) }",
        vec!["2"]
    ),
    iter_seq_yield_false_halts => (
        "package main; import \"fmt\"; func main() { stopped := 0; seq := func(yield func(int) bool) { if !yield(1) { stopped = 1; return }; yield(2) }; for v := range seq { if v == 1 { break } }; fmt.Println(stopped) }",
        vec!["1"]
    ),
    iter_pull2_second_pair => (
        "package main; import \"fmt\"; import \"iter\"; func main() { seq := func(yield func(string, int) bool) { yield(\"a\", 1); yield(\"b\", 2) }; next, stop := iter.Pull2(seq); defer stop(); next(); k, v, ok := next(); fmt.Println(k); fmt.Println(v); fmt.Println(ok) }",
        vec!["b", "2", "true"]
    ),
    iter_slices_all_modify_via_index => (
        "package main; import \"fmt\"; import \"slices\"; func main() { s := []int{1, 2, 3}; for i := range slices.All(s) { s[i] *= 2 }; fmt.Println(s[0]); fmt.Println(s[2]) }",
        vec!["2", "6"]
    ),
    iter_nested_range_slices_all => (
        "package main; import \"fmt\"; import \"slices\"; func main() { outer := [][]int{{1, 2}, {3}}; total := 0; for oi := range slices.All(outer) { for ii := range slices.All(outer[oi]) { total += outer[oi][ii] } }; fmt.Println(total) }",
        vec!["6"]
    ),
    iter_maps_values_after_assignment => (
        "package main; import \"fmt\"; import \"maps\"; func main() { m := map[int]int{1: 1}; m[2] = 2; sum := 0; for v := range maps.Values(m) { sum += v }; fmt.Println(sum) }",
        vec!["3"]
    ),
    iter_pull_stop_before_next => (
        "package main; import \"fmt\"; import \"iter\"; func main() { ran := 0; seq := func(yield func(int) bool) { ran++; yield(1); yield(2) }; next, stop := iter.Pull(seq); stop(); fmt.Println(ran) }",
        vec!["0"]
    ),
    iter_seq_filter_with_yield => (
        "package main; import \"fmt\"; func main() { seq := func(yield func(int) bool) { for i := 1; i <= 5; i++ { if i%2 == 0 { if !yield(i) { return } } } }; evens := 0; for range seq { evens++ }; fmt.Println(evens) }",
        vec!["2"]
    ),
    iter_slices_values_string_map => (
        "package main; import \"fmt\"; import \"slices\"; func main() { m := map[int]string{1: \"go\", 2: \"lang\"}; longest := 0; for v := range slices.Values(m) { if len(v) > longest { longest = len(v) } }; fmt.Println(longest) }",
        vec!["4"]
    ),
    iter_pull2_empty_seq => (
        "package main; import \"fmt\"; import \"iter\"; func main() { seq := func(yield func(int, int) bool) {}; next, stop := iter.Pull2(seq); defer stop(); _, _, ok := next(); fmt.Println(ok) }",
        vec!["false"]
    ),
    iter_slices_all_nil_slice => (
        "package main; import \"fmt\"; import \"slices\"; func main() { var s []int; n := 0; for range slices.All(s) { n++ }; fmt.Println(n) }",
        vec!["0"]
    ),
    iter_maps_values_nil_map => (
        "package main; import \"fmt\"; import \"maps\"; func main() { var m map[int]int; n := 0; for range maps.Values(m) { n++ }; fmt.Println(n) }",
        vec!["0"]
    ),
    iter_seq_first_value_only => (
        "package main; import \"fmt\"; func main() { seq := func(yield func(int) bool) { yield(99); yield(100) }; first := 0; for v := range seq { first = v; break }; fmt.Println(first) }",
        vec!["99"]
    ),
}

go_compile_cases! {
    iter_seq_type_as_func_value => "package main; import \"iter\"; func main() { var seq iter.Seq[int] = func(yield func(int) bool) { yield(1) }; _ = seq }",
    iter_seq2_type_key_value => "package main; import \"iter\"; func main() { var seq iter.Seq2[string, int] = func(yield func(string, int) bool) { yield(\"k\", 1) }; _ = seq }",
    iter_pull_from_slices_all => "package main; import \"iter\"; import \"slices\"; func main() { s := []int{1, 2}; next, stop := iter.Pull(slices.All(s)); defer stop(); _, _ = next() }",
    iter_pull2_from_maps_values => "package main; import \"iter\"; import \"maps\"; func main() { m := map[int]string{1: \"a\"}; next, stop := iter.Pull2(maps.All(m)); defer stop(); _, _, _ = next() }",
    iter_pull_from_custom_seq => "package main; import \"iter\"; func nums() iter.Seq[int] { return func(yield func(int) bool) { yield(1); yield(2) } }; func main() { next, stop := iter.Pull(nums()); defer stop(); _ = next }",
    iter_slices_backward_compile => "package main; import \"slices\"; func main() { s := []int{1, 2, 3}; for i := range slices.Backward(s) { _ = s[i] } }",
    iter_slices_values_struct_map => "package main; import \"slices\"; type P struct { N int }; func main() { m := map[string]P{\"a\": {1}}; for range slices.Values(m) {}",
    iter_maps_all_key_value => "package main; import \"maps\"; func main() { m := map[int]string{1: \"a\"}; for k, v := range maps.All(m) { _, _ = k, v } }",
    iter_seq_nested_pull => "package main; import \"iter\"; func main() { outer := func(yield func(iter.Seq[int]) bool) { yield(func(yield func(int) bool) { yield(1) }) }; for inner := range outer { next, stop := iter.Pull(inner); stop() } }",
    iter_pull_defer_stop_pattern => "package main; import \"iter\"; func main() { seq := func(yield func(int) bool) { yield(1) }; func consume() { next, stop := iter.Pull(seq); defer stop(); _, _ = next() }; consume() }",
    iter_seq2_range_three_pairs => "package main; import \"iter\"; func main() { seq := func(yield func(int, int) bool) { yield(1, 10); yield(2, 20); yield(3, 30) }; for k, v := range seq { _, _ = k, v } }",
    iter_slices_chunk_compile => "package main; import \"slices\"; func main() { s := []int{1, 2, 3, 4}; for c := range slices.Chunk(s, 2) { _ = c } }",
    iter_slices_values_over_keys_via_maps => "package main; import \"maps\"; import \"slices\"; func main() { m := map[string]int{\"a\": 1}; for range slices.Values(m) {}",
    iter_seq_yield_returns_bool => "package main; func main() { seq := func(yield func(int) bool) bool { return yield(1) }; _ = seq }",
    iter_pull2_manual_iteration => "package main; import \"iter\"; func main() { seq := func(yield func(rune, rune) bool) { yield('a', 'b') }; next, stop := iter.Pull2(seq); defer stop(); for { _, _, ok := next(); if !ok { break } } }",
    iter_maps_keys_iterates => "package main; import \"maps\"; func main() { m := map[int]int{1: 1, 2: 2}; for k := range maps.Keys(m) { _ = k } }",
    iter_seq_closure_capture => "package main; import \"iter\"; func main() { base := 10; seq := func(yield func(int) bool) { yield(base) }; for v := range seq { _ = v } }",
    iter_pull_from_empty_seq => "package main; import \"iter\"; func main() { seq := func(yield func(int) bool) {}; next, stop := iter.Pull(seq); defer stop(); _, _ = next() }",
    iter_slices_all_over_bytes => "package main; import \"slices\"; func main() { b := []byte{'a', 'b'}; for i := range slices.All(b) { _ = b[i] } }",
    iter_seq2_string_int_map_range => "package main; import \"maps\"; func main() { for k, v := range maps.All(map[string]int{\"x\": 1}) { _, _ = k, v } }",
}
