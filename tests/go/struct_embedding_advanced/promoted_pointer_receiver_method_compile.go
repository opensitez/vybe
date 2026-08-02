// vybe-test: go/struct_embedding_advanced/promoted_pointer_receiver_method_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type inner struct { n int }
func (i *inner) bump() { i.n++ }
type outer struct { inner }
func main() { var value outer
value.bump() }
