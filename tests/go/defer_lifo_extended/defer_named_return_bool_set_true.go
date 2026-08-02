// vybe-test: go/defer_lifo_extended/defer_named_return_bool_set_true
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (ok bool) { defer func() { ok = true }()
return false }
func main() { fmt.Println(work()) }
