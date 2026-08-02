// vybe-test: go/unsafe_size_align_extended/unsafe_sizeof_struct_three_fields
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs

package main
import "fmt"
import "unsafe"
type S struct { x int16
y int16
z int32 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(unsafe.Sizeof(S{})), "8") }
