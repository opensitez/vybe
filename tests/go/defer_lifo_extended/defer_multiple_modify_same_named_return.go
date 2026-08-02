// vybe-test: go/defer_lifo_extended/defer_multiple_modify_same_named_return
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (n int) { defer func() { n += 1 }()
defer func() { n += 10 }()
return 0 }
func main() { fmt.Println(work()) }
