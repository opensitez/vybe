// vybe-test: go/defer_panic_variants/defer_in_loop_with_break_registers_three
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func main() { for i := 0; i < 5; i++ { defer fmt.Println(i)
if i == 2 { break } } }
