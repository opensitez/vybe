// vybe-test: go/interface_nil_comparable/interface_compare_after_type_assertion_compile
// origin: languages/go/tests/go/test_interface_nil_comparable.rs
// vybe-test-mode: compile

package main
func main() { var value interface{} = 2
number, ok := value.(int)
if ok { _ = (number == 2) } }
