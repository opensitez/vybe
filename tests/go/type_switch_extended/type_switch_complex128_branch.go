// vybe-test: go/type_switch_extended/type_switch_complex128_branch
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case complex128: fmt.Println("cx") default: fmt.Println("other") } }
func main() { tag(complex(2, 3)) }
