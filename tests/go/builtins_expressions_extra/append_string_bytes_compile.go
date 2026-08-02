// vybe-test: go/builtins_expressions_extra/append_string_bytes_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { dst := []byte{'a'}
dst = append(dst, []byte("bc")...)
_ = dst }
