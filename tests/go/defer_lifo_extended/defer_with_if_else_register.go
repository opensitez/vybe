// vybe-test: go/defer_lifo_extended/defer_with_if_else_register
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { ok := true
if ok { defer fmt.Println("yes") } else { defer fmt.Println("no") } }
