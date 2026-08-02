// vybe-test: go/for_range_extended/range_int_labeled_continue_skip
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
outer: for i := range 5 { if i%2 == 0 { continue outer }
total += i }
fmt.Println(total) }
