// vybe-test: go/nil_zero_semantics_extra/zero_value_array_compare_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var a [2]int
var b [2]int
_ = (a == b) }
