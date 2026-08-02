// vybe-test: go/go_statement_compile/go_closure_capture_slice_element_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
func main() { s := []int{1, 2}
go func() { _ = s[0] }() }
