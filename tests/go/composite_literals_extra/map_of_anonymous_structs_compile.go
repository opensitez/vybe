// vybe-test: go/composite_literals_extra/map_of_anonymous_structs_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]struct { n int }{"x": {n: 1}}
_ = values }
