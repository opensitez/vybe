// vybe-test: go/defer_lifo_extended/defer_named_return_increment_twice
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (n int) { defer func() { n++ }()
defer func() { n++ }()
return 0 }
func main() { fmt.Println(work()) }
