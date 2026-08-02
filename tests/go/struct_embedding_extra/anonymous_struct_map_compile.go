// vybe-test: go/struct_embedding_extra/anonymous_struct_map_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]struct { n int }{"a": {n: 1}}
_ = values }
