// vybe-test: go/blank_identifier_extended/blank_import_with_named_fmt_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
import ( "fmt"; _ "strings" )
func main() { _ = fmt.Sprint(1) }
