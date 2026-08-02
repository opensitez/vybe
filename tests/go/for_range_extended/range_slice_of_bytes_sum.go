// vybe-test: go/for_range_extended/range_slice_of_bytes_sum
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for _, b := range []byte{'a', 'b', 'c'} { total += int(b) }
fmt.Println(total) }
