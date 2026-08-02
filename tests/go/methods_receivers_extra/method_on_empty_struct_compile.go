// vybe-test: go/methods_receivers_extra/method_on_empty_struct_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type marker struct{}
func (marker) ok() bool { return true }
func main() { _ = marker{}.ok() }
