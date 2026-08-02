// vybe-test: go/defer_panic_recover_extra/defer_order_with_named_functions_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func one() { fmt.Println(1) }
func two() { fmt.Println(2) }
func main() { defer one()
defer two()
}
