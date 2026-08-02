// vybe-test: go/init_blank_import/init_with_switch_statement_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
var tag string
func init() { switch 2 { case 2: tag = "hit" default: tag = "miss" } }
func main() { _ = tag }
