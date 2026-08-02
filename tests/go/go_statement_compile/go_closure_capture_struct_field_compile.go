// vybe-test: go/go_statement_compile/go_closure_capture_struct_field_compile
// origin: languages/go/tests/go/test_go_statement_compile.rs
// vybe-test-mode: compile

package main
type box struct { n int }
func main() { b := box{n: 2}
go func() { _ = b.n }() }
