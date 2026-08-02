// vybe-test: go/generics_constraints_extended/generic_tilde_struct_slice_len
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Ints []int
func Total[S ~[]int](s S) int { sum := 0
for _, v := range s { sum += v }
return sum }
func main() { fmt.Println(Total(Ints{1, 2, 3})) }
