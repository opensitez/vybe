// vybe-test: go/init_function_order/init_after_blank_import_strings_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
import _ "strings"
var seeded int
func init() { seeded = 1 }
func init() { seeded++ }
func main() { _ = seeded }
