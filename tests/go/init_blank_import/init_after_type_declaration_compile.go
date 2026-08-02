// vybe-test: go/init_blank_import/init_after_type_declaration_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
type score int
var high score
func init() { high = score(99) }
func main() { _ = high }
