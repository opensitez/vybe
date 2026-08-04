// vybe-test: go/concurrency_dummy/go_routine_closure
// origin: languages/go/tests/go/test_concurrency_dummy.rs
// vybe-test-mode: compile

package main
func main() { go func() { }() }
