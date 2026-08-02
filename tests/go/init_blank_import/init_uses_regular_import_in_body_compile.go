// vybe-test: go/init_blank_import/init_uses_regular_import_in_body_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
import "fmt"
var label string
func init() { label = fmt.Sprintf("%s-%d", "init", 3) }
func main() { _ = label }
