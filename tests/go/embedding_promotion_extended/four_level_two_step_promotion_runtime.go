// vybe-test: go/embedding_promotion_extended/four_level_two_step_promotion_runtime
// origin: languages/go/tests/go/test_embedding_promotion_extended.rs

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

func main() { v := a{b: b{c: c{d: d{n: 13}}}}
__check(fmt.Sprint(v.n), "13") }
