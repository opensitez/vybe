// vybe-test: go/lang_functions_returns/defer_call_args_evaluated_early
// origin: languages/go/tests/go/test_lang_functions_returns.rs
// vybe-test-mode: compile

package main
func main() { defer func(int) {}(func() int { return 1 }()) }
