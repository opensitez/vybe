// vybe-test: go/types_advanced/slice_of_interface
// origin: languages/go/tests/go/test_types_advanced.rs
// vybe-test-mode: compile

package main
func main() { s := []interface{}{1, "str", true}
_ = s }
