// vybe-test: go/defer_panic_recover_extra/defer_modify_slice_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func main() { values := []int{1}
func() { defer func() { values[0] = 9 }() }()
fmt.Println(values[0])
}
