// vybe-test: go/lang_declarations_types/three_index_slice
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{0,1,2,3,4}
t := s[1:3:4]
__check(fmt.Sprint(len(t)) + " " + fmt.Sprint(cap(t)), "2 3") }
