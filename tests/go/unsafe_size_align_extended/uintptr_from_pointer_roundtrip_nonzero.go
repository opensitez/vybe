// vybe-test: go/unsafe_size_align_extended/uintptr_from_pointer_roundtrip_nonzero
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs

package main
import "fmt"
import "unsafe"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var x int = 3
u := uintptr(unsafe.Pointer(&x))
p := unsafe.Pointer(u)
__check(fmt.Sprint(p != nil), "true") }
