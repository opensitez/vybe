// vybe-test: go/lang_declarations_types/slice_not_comparable_compile
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
func main() { _ = []int{1} == []int{1} }
