// vybe-test: go/for_range_extended/range_slice_modify_value_via_index
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { nums := []int{1, 2, 3}
for i := range nums { nums[i] *= 2 }
fmt.Println(nums[1]) }
