// vybe-test: go/unsafe_size_align_extended/unsafe_sizeof_nested_struct
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs

package main
import "fmt"
import "unsafe"
type Inner struct { v int32 }
type Outer struct { i Inner
flag bool }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(unsafe.Sizeof(Outer{})), "8") }
