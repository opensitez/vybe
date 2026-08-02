// vybe-test: go/blank_identifier_extended/blank_discard_map_delete_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { m := map[string]int{"a": 1}
delete(m, "a")
_, ok := m["a"]
_ = ok }
