// vybe-test: go/for_range_extended/range_over_rune_slice_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { runes := []rune("xy")
for i, r := range runes { _, _ = i, r } }
