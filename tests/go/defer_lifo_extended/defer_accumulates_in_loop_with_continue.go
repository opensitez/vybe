// vybe-test: go/defer_lifo_extended/defer_accumulates_in_loop_with_continue
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { for i := 0; i < 3; i++ { defer fmt.Println(i)
if i == 0 { continue } } }
