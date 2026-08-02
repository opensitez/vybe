// vybe-test: go/struct_embedding_advanced/four_level_field_promotion_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type d struct { n int }
type c struct { d }
type b struct { c }
type a struct { b }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := a{b: b{c: c{d: d{n: 13}}}}
__check(fmt.Sprint(value.n), "13")
}
