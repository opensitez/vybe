// vybe-test: go/pointer_receivers_advanced/nil_receiver_recover_wrapper_compile
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs
// vybe-test-mode: compile

package main
type widget struct { size int }
func (w *widget) read() int { return w.size }
func safe(v *widget) { defer func() { recover() }()
_ = v.read() }
func main() { var value *widget
safe(value) }
