// vybe-test: go/method_sets_pointer_value/value_variable_pointer_only_method_expression_compile_fail
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile-fail

package main
type cell struct { n int }
func (c *cell) set(v int) { c.n = v }
func main() { v := cell{}
fn := cell.set
_ = fn }
