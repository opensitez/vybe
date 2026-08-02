// vybe-test: go/nil_zero_semantics_extra/zero_value_recursive_struct_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
type node struct { next *node }
func main() { var n node
_ = n }
