// vybe-test: go/unsafe_size_align_extended/no_arithmetic_string_data_len
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
func main() { s := "go"
ptr := unsafe.StringData(s)
_ = ptr
_ = len(s) }
