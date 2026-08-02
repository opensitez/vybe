// vybe-test: go/struct_embedding_advanced/triple_nested_embedding_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type leaf struct { value int }
type branch struct { leaf }
type trunk struct { branch }
func main() { var value trunk
_ = value.value }
