// vybe-test: go/defer_panic_recover_extra/defer_local_cleanup_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func main() { done := 0
func() { defer func() { done = 3 }() }()
fmt.Println(done)
}
