// vybe-test: go/composite_literals_extra/composite_literal_assigned_to_interface_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
type any interface{}
func main() { var v any = struct { n int }{n: 7}
_ = v }
