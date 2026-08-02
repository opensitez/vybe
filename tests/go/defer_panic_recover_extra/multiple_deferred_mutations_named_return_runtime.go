// vybe-test: go/defer_panic_recover_extra/multiple_deferred_mutations_named_return_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build() (result int) { defer func() { result += 3 }()
defer func() { result *= 2 }()
result = 4
return }
func main() { fmt.Println(build())
}
