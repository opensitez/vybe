// vybe-test: go/method_sets_pointer_value/interface_field_with_embedded_implementer_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type doer interface { doWork() }
type impl struct{}
func (impl) doWork() {}
type holder struct { doer }
func main() { _ = holder{doer: impl{}} }
