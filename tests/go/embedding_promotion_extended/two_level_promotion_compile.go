// vybe-test: go/embedding_promotion_extended/two_level_promotion_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type leaf struct { v int }
type branch struct { leaf }
type trunk struct { branch }
func main() { _ = trunk{}.v }
