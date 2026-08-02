// vybe-test: go/interfaces_patterns_extra/interface_assignment_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var left interface{} = 1
var right interface{} = left
_ = right }
