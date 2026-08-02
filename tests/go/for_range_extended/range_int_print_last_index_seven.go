// vybe-test: go/for_range_extended/range_int_print_last_index_seven
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { last := -1
for i := range 7 { last = i }
fmt.Println(last) }
