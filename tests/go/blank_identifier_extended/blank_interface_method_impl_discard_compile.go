// vybe-test: go/blank_identifier_extended/blank_interface_method_impl_discard_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
type Closer interface { Close() error }
type resource struct{}
func (r resource) Close() error { return nil }
func main() { var c Closer = resource{}
_ = c }
