// vybe-test: go/defer_lifo_extended/defer_on_label_statement_in_loop
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { for i := 0; i < 2; i++ { defer fmt.Println(i)
if i == 1 { break } } }
