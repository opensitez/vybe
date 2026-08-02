// vybe-test: go/embedding_promotion_extended/shadowed_field_both_accessible_compile
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs
// vybe-test-mode: compile

package main
type base struct { id int }
type derived struct { base
id int }
func main() { var d derived
_ = d.id
_ = d.base.id }
