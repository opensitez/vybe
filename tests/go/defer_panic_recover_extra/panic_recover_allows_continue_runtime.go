// vybe-test: go/defer_panic_recover_extra/panic_recover_allows_continue_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func safe() { defer func() { recover() }()
panic("x") }
func main() { safe()
fmt.Println(1)
}
