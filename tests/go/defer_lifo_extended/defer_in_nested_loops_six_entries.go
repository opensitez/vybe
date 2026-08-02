// vybe-test: go/defer_lifo_extended/defer_in_nested_loops_six_entries
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func main() { for i := 0; i < 2; i++ { for j := 0; j < 3; j++ { defer fmt.Println(i*10+j) } } }
