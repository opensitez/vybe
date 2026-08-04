// vybe-test: go/composite_literals_extra/pointer_to_array_literal_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
func main() { values := &[3]int{1, 2, 3}
_ = values }
