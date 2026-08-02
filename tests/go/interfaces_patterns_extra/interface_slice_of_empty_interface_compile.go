// vybe-test: go/interfaces_patterns_extra/interface_slice_of_empty_interface_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []interface{}{1, "go"}
_ = values }
