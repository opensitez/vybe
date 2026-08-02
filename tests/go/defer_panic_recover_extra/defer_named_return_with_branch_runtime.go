// vybe-test: go/defer_panic_recover_extra/defer_named_return_with_branch_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build(flag bool) (result int) { defer func() { result++ }()
if flag { return 5 }
return 2 }
func main() { fmt.Println(build(true))
}
