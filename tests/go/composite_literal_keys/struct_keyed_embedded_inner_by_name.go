// vybe-test: go/composite_literal_keys/struct_keyed_embedded_inner_by_name
// origin: languages/go/tests/go/test_composite_literal_keys.rs

package main
import "fmt"
type inner struct { value int }
type outer struct { inner }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { o := outer{inner: inner{value: 42}}
__check(fmt.Sprint(o.value), "42")
}
