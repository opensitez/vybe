// vybe-test: go/range_over_int/range_int_last_index
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { last := -1
for i := range 5 { last = i }
fmt.Println(last) }
