// vybe-test: go/for_range_extended/range_slice_reassign_loop_var
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { sum := 0
for _, v := range []int{1, 2, 3} { v = v * 10
sum += v }
fmt.Println(sum) }
