// vybe-test: go/interface_nil_comparable/interface_equality_compile
// origin: languages/go/tests/go/test_interface_nil_comparable.rs
// vybe-test-mode: compile

package main
func main() { var left interface{} = 1
var right interface{} = 1
_ = (left == right) }
