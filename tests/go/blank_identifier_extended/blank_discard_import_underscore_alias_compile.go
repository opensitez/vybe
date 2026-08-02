// vybe-test: go/blank_identifier_extended/blank_discard_import_underscore_alias_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
import f "fmt"
import _ "strings"
func main() { _ = f.Sprint("x") }
