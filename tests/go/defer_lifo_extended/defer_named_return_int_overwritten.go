// vybe-test: go/defer_lifo_extended/defer_named_return_int_overwritten
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (n int) { defer func() { n = 20 }()
return 5 }
func main() { fmt.Println(work()) }
