// vybe-test: go/defer_lifo_extended/defer_method_on_nil_pointer_compile
// origin: languages/go/tests/go/test_defer_lifo_extended.rs
// vybe-test-mode: compile

package main
type T struct{}
func (t *T) f() {}
func main() { var t *T
defer t.f() }
