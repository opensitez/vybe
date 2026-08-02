// vybe-test: go/type_switch_extended/type_switch_on_non_interface_compile
// origin: languages/go/tests/go/test_type_switch_extended.rs
// vybe-test-mode: compile

package main
func main() { x := 1
switch x.(type) { case int: _ = x } }
