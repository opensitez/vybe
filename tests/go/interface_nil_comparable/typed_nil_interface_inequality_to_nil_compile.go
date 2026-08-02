// vybe-test: go/interface_nil_comparable/typed_nil_interface_inequality_to_nil_compile
// origin: languages/go/tests/go/test_interface_nil_comparable.rs
// vybe-test-mode: compile

package main
func main() { var p *int
var value interface{} = p
_ = (value != nil) }
