// vybe-test: go/defer_lifo_extended/defer_for_loop_break_registers_partial
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { for i := 0; i < 5; i++ { defer fmt.Println(i)
if i == 1 { break } } }
