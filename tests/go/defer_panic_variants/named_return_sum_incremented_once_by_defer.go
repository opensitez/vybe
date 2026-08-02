// vybe-test: go/defer_panic_variants/named_return_sum_incremented_once_by_defer
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func total() (sum int) { defer func() { sum = sum + 10 }()
return 7 }
func main() { fmt.Println(total())
}
