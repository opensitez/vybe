// vybe-test: go/defer_panic_variants/named_return_bare_return_scaled_by_defer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func scale() (n int) { defer func() { n = n * 3 }()
n = 4
return }
func main() { fmt.Println(scale())
}
