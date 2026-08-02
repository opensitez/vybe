// vybe-test: go/unsafe_size_align_extended/unsafe_offsetof_string_field
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs

package main
import "fmt"
import "unsafe"
type S struct { name string
n int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(unsafe.Offsetof(S{}.n)), "16") }
