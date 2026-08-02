// vybe-test: go/unsafe_size_align_extended/uintptr_from_string_data
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { s := "a"
_ = uintptr(unsafe.Pointer(unsafe.StringData(s))) }
