// vybe-test: go/methods_receivers_extra/method_with_slice_field_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type bag struct { values []int }
func (b bag) count() int { return len(b.values) }
func main() { _ = bag{}.count() }
