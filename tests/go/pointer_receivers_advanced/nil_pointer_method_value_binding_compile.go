// vybe-test: go/pointer_receivers_advanced/nil_pointer_method_value_binding_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type widget struct { size int }
func (w *widget) grow() { w.size++ }
func main() { var value *widget
fn := value.grow
fn() }
