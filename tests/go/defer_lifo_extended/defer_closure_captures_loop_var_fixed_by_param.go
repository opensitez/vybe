// vybe-test: go/defer_lifo_extended/defer_closure_captures_loop_var_fixed_by_param
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { for i := 0; i < 3; i++ { defer func(n int) { fmt.Println(n) }(i) }
}
