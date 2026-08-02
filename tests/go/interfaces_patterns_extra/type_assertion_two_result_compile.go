// vybe-test: go/interfaces_patterns_extra/type_assertion_two_result_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var value interface{} = 1
number, ok := value.(int)
_, _ = number, ok }
