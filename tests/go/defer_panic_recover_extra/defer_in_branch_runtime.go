// vybe-test: go/defer_panic_recover_extra/defer_in_branch_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func main() { if true { defer fmt.Println(2) }
fmt.Println(1)
}
