// vybe-test: go/defer_panic_variants/defer_in_loop_accumulates_counter_on_exit
// origin: languages/go/tests/go/test_defer_panic_variants.rs

package main
import "fmt"
func main() { total := 0
for i := 1; i <= 3; i++ { defer func() { total = total + i }() }
fmt.Println(total) }
