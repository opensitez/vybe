// vybe-test: go/for_range_extended/range_int_decrement_pattern
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { acc := 0
for i := range 4 { acc = acc*10 + (3 - i) }
fmt.Println(acc) }
