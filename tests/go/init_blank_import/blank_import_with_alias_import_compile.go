// vybe-test: go/init_blank_import/blank_import_with_alias_import_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
import f "fmt"
import _ "strings"
func main() { _ = f.Sprint("ok") }
