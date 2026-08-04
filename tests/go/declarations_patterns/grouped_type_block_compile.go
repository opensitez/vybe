// vybe-test: go/declarations_patterns/grouped_type_block_compile
// origin: languages/go/tests/go/test_declarations_patterns.rs
// vybe-test-mode: compile

package main
type ( Score int; Label string )
func main() { var s Score = 3
var l Label = "ok"
_, _ = s, l }
