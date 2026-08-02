// vybe-test: go/lang_declarations_types/map_not_comparable_compile
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
func main() { _ = map[int]int{} == map[int]int{} }
