// vybe-test: go/lang_generics_semantics/generic_map_make
// origin: languages/go/tests/go/test_lang_generics_semantics.rs
// vybe-test-mode: compile

package main
func MakeMap[K comparable, V any]() map[K]V { return make(map[K]V) }
func main() { _ = MakeMap[string, int]() }
