// vybe-test: go/methods_receivers_extra/method_with_map_field_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type bag struct { values map[string]int }
func (b bag) size() int { return len(b.values) }
func main() { _ = bag{}.size() }
