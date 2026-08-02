// vybe-test: go/lang_declarations_types/dot_import_compile
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
import . "fmt"
func main() { Println(1) }
