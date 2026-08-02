// vybe-test: go/defer_lifo_extended/defer_named_return_bare_return_uses_updated
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (n int) { defer func() { n = n + 1 }()
n = 4
return }
func main() { fmt.Println(work()) }
