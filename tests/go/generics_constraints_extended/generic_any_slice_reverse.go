// vybe-test: go/generics_constraints_extended/generic_any_slice_reverse
// origin: languages/go/tests/go/test_generics_constraints_extended.rs

package main
import "fmt"
func Reverse[T any](s []T) { for i, j := 0, len(s)-1; i < j; i, j = i+1, j-1 { s[i], s[j] = s[j], s[i] } }
func main() { a := []int{1, 2, 3}
Reverse(a)
fmt.Println(a[0])
fmt.Println(a[2]) }
