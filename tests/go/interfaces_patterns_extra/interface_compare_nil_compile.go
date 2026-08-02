// vybe-test: go/interfaces_patterns_extra/interface_compare_nil_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var value interface{}
_ = (value == nil) }
