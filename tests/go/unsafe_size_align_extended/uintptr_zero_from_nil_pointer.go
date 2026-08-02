// vybe-test: go/unsafe_size_align_extended/uintptr_zero_from_nil_pointer
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

func main() { var p *int
__check(fmt.Sprint(uintptr(unsafe.Pointer(p)) == 0), "true") }
