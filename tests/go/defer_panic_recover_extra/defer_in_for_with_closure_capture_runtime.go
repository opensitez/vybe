// vybe-test: go/defer_panic_recover_extra/defer_in_for_with_closure_capture_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func main() { for i := 0; i < 2; i++ { value := i
defer func() { fmt.Println(value) }() } }
