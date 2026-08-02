// vybe-test: go/unsafe_size_align_extended/unsafe_sizeof_struct_no_padding
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs

package main
import "fmt"
import "unsafe"
type S struct { a int32
b int32 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(unsafe.Sizeof(S{})), "8") }
