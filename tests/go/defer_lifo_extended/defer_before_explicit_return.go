// vybe-test: go/defer_lifo_extended/defer_before_explicit_return
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() int { defer fmt.Println("defer")
fmt.Println("body")
return 1 }
func main() { fmt.Println(work()) }
