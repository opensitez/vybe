// vybe-test: go/builtins_expressions_extra/recover_in_deferred_closure_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { defer func(message string) { _ = recover()
_ = message }("done")
panic("boom") }
