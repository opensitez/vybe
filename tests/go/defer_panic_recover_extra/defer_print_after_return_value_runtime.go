// vybe-test: go/defer_panic_recover_extra/defer_print_after_return_value_runtime
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs

package main
import "fmt"
func build() int { defer fmt.Println("later")
return 4 }
func main() { fmt.Println(build())
}
