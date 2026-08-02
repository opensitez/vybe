// vybe-test: go/init_blank_import/grouped_blank_imports_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
import ( _ "strings"; _ "math" )
func main() {}
