// vybe-test: go/struct_embedding_extra/anonymous_struct_slice_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []struct { n int }{{n: 1}}
_ = values }
