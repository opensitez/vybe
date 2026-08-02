// vybe-test: go/defer_lifo_extended/defer_named_return_float64_scaled
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (f float64) { defer func() { f = f * 2 }()
return 3.5 }
func main() { fmt.Println(f == 7.0) }
