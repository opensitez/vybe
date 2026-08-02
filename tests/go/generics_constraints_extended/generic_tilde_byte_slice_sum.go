// vybe-test: go/generics_constraints_extended/generic_tilde_byte_slice_sum
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
type Bytes []byte
func SumBytes[B ~[]byte](b B) int { s := 0
for _, c := range b { s += int(c) }
return s }
func main() { fmt.Println(SumBytes(Bytes{'a', 'b'})) }
