// vybe-test: go/composite_literal_keys/nested_anonymous_struct_all_keyed_compile
// origin: languages/go/tests/go/test_composite_literal_keys.rs
// vybe-test-mode: compile

package main
func main() { _ = struct { outer struct { n int } }{outer: struct { n int }{n: 5}} }
