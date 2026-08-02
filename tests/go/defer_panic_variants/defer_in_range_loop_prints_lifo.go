// vybe-test: go/defer_panic_variants/defer_in_range_loop_prints_lifo
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func main() { for _, value := range []int{10, 20, 30} { defer fmt.Println(value) } }
