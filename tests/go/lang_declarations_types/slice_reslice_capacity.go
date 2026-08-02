// vybe-test: go/lang_declarations_types/slice_reslice_capacity
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := make([]int, 2, 4)
t := s[:3]
__check(fmt.Sprint(cap(t)), "4") }
