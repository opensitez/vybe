// vybe-test: go/error_handling/multiple_panics_in_defer
// origin: languages/go/tests/go/test_error_handling.rs
// vybe-test-mode: compile

package main
func main() { defer func() { recover() }()
panic("1") }
