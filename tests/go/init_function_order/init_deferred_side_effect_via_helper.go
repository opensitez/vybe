// vybe-test: go/init_function_order/init_deferred_side_effect_via_helper
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var order string
func mark(c string) { order += c }
func init() { mark("1")
defer mark("d1") }
func init() { mark("2")
defer mark("d2") }
func main() { fmt.Println(order) }
