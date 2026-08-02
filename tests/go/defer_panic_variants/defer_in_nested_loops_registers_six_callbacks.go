// vybe-test: go/defer_panic_variants/defer_in_nested_loops_registers_six_callbacks
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func main() { for i := 0; i < 2; i++ { for j := 0; j < 3; j++ { defer fmt.Println(i*10 + j) } } }
