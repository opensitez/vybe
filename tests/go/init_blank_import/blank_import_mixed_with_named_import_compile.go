// vybe-test: go/init_blank_import/blank_import_mixed_with_named_import_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
import ( "fmt"; _ "strings" )
func main() { _ = fmt.Sprintf("%d", 1) }
