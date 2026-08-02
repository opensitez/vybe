// vybe-test: go/defer_panic_recover_extra/defer_before_multiple_returns_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build(flag bool) int { defer fmt.Println("done")
if flag { return 1 }
return 2 }
func main() { fmt.Println(build(false))
}
