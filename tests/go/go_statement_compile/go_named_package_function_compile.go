// vybe-test: go/go_statement_compile/go_named_package_function_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func tick() {}
func main() { go tick() }
