// vybe-test: go/defer_lifo_extended/defer_after_short_var_in_if
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { if x := 1; x > 0 { defer fmt.Println(x)
}
}
