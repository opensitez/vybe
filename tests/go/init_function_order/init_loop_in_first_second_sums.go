// vybe-test: go/init_function_order/init_loop_in_first_second_sums
// origin: languages/go/tests/go/test_init_function_order.rs

package main
import "fmt"
var seed int
var total int
func init() { for i := 0; i < 3; i++ { seed += i } }
func init() { total = seed + 10 }
func main() { fmt.Println(total) }
