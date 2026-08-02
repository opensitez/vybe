// vybe-test: go/blank_identifier_extended/blank_discard_slice_index
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{10, 20, 30}
_, last := s[0], s[2]
__check(fmt.Sprint(last), "30") }
