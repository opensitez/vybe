// vybe-test: go/type_aliases/alias_map_type_in_struct_field
// origin: languages/go/tests/go/test_type_aliases.rs
// vybe-test-mode: compile

package main
type Dict = map[string]int
type holder struct { data Dict }
func main() { _ = holder{data: Dict{"k": 1}} }
