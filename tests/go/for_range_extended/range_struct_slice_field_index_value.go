// vybe-test: go/for_range_extended/range_struct_slice_field_index_value
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
type bag struct { items []int }
func main() { b := bag{items: []int{4, 5}}
total := 0
for i, v := range b.items { total += i + v }
fmt.Println(total) }
