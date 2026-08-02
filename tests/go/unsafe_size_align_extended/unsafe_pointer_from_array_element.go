// vybe-test: go/unsafe_size_align_extended/unsafe_pointer_from_array_element
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

func main() { arr := [2]int{1, 2}
p := unsafe.Pointer(&arr[1])
__check(fmt.Sprint(p != nil), "true") }
