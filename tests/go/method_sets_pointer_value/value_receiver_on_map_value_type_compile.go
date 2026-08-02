// vybe-test: go/method_sets_pointer_value/value_receiver_on_map_value_type_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type key struct { s string }
func (k key) hash() int { return len(k.s) }
func main() { _ = key{s: "a"}.hash() }
