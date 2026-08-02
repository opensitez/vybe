// vybe-test: go/interfaces_patterns_extra/interface_method_returning_interface_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
type builder interface { build() interface{} }
func main() {}
