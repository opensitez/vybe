// vybe-test: go/nil_zero_semantics_extra/nil_pointer_in_method_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
type counter struct{}
func (c *counter) ok() bool { return c == nil }
func main() { var c *counter
_ = c.ok() }
