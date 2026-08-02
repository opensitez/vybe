// vybe-test: go/blank_identifier_extended/blank_anonymous_struct_literal_embed_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
func main() { type outer struct { _ struct{}
n int }
o := outer{n: 2}
_ = o.n }
