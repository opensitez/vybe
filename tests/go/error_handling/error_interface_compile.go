// vybe-test: go/error_handling/error_interface_compile
// origin: languages/go/tests/go/test_error_handling.rs
// vybe-test-mode: compile

package main
type error interface { Error() string }
func main() {}
