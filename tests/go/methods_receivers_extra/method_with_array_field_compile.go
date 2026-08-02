// vybe-test: go/methods_receivers_extra/method_with_array_field_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type bag struct { values [2]int }
func (b bag) first() int { return b.values[0] }
func main() { _ = bag{}.first() }
