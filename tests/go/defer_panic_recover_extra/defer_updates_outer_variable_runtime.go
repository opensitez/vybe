// vybe-test: go/defer_panic_recover_extra/defer_updates_outer_variable_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func main() { total := 1
func() { defer func() { total = 8 }() }()
fmt.Println(total)
}
