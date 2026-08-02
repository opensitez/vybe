// vybe-test: go/defer_panic_recover_extra/defer_named_return_add_two_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build() (result int) { defer func() { result += 2 }()
return 3 }
func main() { fmt.Println(build())
}
