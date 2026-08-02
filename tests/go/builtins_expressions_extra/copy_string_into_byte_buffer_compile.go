// vybe-test: go/builtins_expressions_extra/copy_string_into_byte_buffer_compile
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs
// vybe-test-mode: compile

package main
func main() { dst := make([]byte, 4)
_ = copy(dst, "go") }
