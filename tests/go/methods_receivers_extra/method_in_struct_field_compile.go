// vybe-test: go/methods_receivers_extra/method_in_struct_field_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func (c counter) total() int { return c.n }
type holder struct { fn func() int }
func main() { value := counter{}
_ = holder{fn: value.total} }
