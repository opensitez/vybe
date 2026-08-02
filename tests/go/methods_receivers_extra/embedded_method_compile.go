// vybe-test: go/methods_receivers_extra/embedded_method_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type inner struct{}
func (inner) label() string { return "ok" }
type outer struct { inner }
func main() { var value outer
_ = value.inner.label() }
