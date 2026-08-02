// vybe-test: go/embedding_promotion_extended/deep_two_level_pointer_method_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type leaf struct { n int }
func (l *leaf) inc() {}
type branch struct { leaf }
type trunk struct { branch }
func main() { var t trunk
t.inc() }
