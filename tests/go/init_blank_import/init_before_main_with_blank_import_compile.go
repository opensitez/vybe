// vybe-test: go/init_blank_import/init_before_main_with_blank_import_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
import _ "strings"
var seeded int
func init() { seeded = 5 }
func main() { _ = seeded }
