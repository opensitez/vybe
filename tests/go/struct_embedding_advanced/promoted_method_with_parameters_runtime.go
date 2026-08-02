// vybe-test: go/struct_embedding_advanced/promoted_method_with_parameters_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct { base int }
func (i inner) add(delta int) int { return i.base + delta }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := outer{inner: inner{base: 3}}
__check(fmt.Sprint(value.add(5)), "8")
}
