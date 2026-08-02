// vybe-test: go/struct_embedding_advanced/nested_pointer_middle_embedding_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type inner struct { count int }
type middle struct { *inner }
type outer struct { middle }
func main() { var value outer
_ = value.count }
