// vybe-test: go/slice_copy_clear/copy_string_to_byte_slice
// origin: languages/go/tests/go/test_slice_copy_clear.rs
// vybe-test-mode: compile

package main
func main() { dst := make([]byte, 3)
_ = copy(dst, "abc") }
