// vybe-test: go/defer_panic_recover_extra/defer_with_argument_copy_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func show(v int) { fmt.Println(v) }
func main() { value := 2
defer show(value)
value = 9
}
