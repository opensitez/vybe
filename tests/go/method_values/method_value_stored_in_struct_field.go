// vybe-test: go/method_values/method_value_stored_in_struct_field
// origin: languages/go/tests/go/test_method_values.rs
// vybe-test-mode: compile

package main
type fn func() int
type holder struct { call fn }
func (c counter) val() int { return 1 }
type counter struct{}
func main() { _ = holder{} }
