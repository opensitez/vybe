// vybe-test: go/init_blank_import/init_with_nested_block_scope_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
var depth int
func init() { { depth = 1
{ depth = depth + 1 } } }
func main() { _ = depth }
