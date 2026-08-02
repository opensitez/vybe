// vybe-test: go/init_function_order/init_uses_imported_fmt_sprint_compile
// origin: languages/go/tests/go/test_init_function_order.rs
// vybe-test-mode: compile

package main
import "fmt"
var s string
func init() { s = fmt.Sprint(1, 2) }
func main() { _ = s }
