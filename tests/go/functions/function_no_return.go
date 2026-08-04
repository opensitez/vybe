// vybe-test: go/functions/function_no_return
// origin: languages/go/tests/go/test_functions.rs
// vybe-test-mode: compile

package main
func printMsg(s string) { _ = s } func main() { printMsg("hello")
}
