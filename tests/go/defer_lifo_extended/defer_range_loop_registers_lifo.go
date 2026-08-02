// vybe-test: go/defer_lifo_extended/defer_range_loop_registers_lifo
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { for _, v := range []int{10, 20} { defer fmt.Println(v) } }
