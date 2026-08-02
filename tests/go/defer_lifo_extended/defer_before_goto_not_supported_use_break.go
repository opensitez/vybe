// vybe-test: go/defer_lifo_extended/defer_before_goto_not_supported_use_break
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { loop: for i := 0; i < 2; i++ { defer fmt.Println(i)
if i == 1 { break loop } } }
