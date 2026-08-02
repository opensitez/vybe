// vybe-test: go/defer_panic_recover_extra/defer_method_call_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
type counter struct{}
func (counter) clean() {}
func main() { value := counter{}
defer value.clean() }
