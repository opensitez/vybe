// vybe-test: go/types_advanced/nested_type_declarations
// origin: languages/go/tests/go/test_types_advanced.rs
// vybe-test-mode: compile

package main
type A struct { B struct { C int } }
func main() {}
