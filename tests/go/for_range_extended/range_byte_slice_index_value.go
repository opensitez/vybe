// vybe-test: go/for_range_extended/range_byte_slice_index_value
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { last := byte(0)
for i, b := range []byte{10, 20} { if i == 1 { last = b } }
fmt.Println(int(last)) }
