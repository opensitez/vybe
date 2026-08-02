// vybe-test: go/defer_panic_variants/defer_lifo_mixed_named_funcs_and_literals
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func mark(label string) { fmt.Println(label) }
func main() { defer mark("alpha")
defer func() { fmt.Println("beta") }()
defer mark("gamma")
}
