// vybe-test: go/blank_identifier_extended/blank_discard_method_call_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
type T struct{}
func (t T) M() int { return 1 }
func main() { _ = T{}.M() }
