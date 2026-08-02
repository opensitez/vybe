// vybe-test: go/method_sets_pointer_value/two_level_embedded_method_promotion_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type leaf struct{}
func (leaf) f() {}
type branch struct { leaf }
type trunk struct { branch }
func main() { trunk{}.f() }
