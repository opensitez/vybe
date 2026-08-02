// vybe-test: go/struct_embedding_advanced/address_of_promoted_field_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type inner struct { count int }
type outer struct { inner }
func main() { var value outer
ptr := &value.count
_ = ptr }
