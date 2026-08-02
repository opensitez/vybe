// vybe-test: go/strings_unicode_extra/byte_slice_to_string_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []byte{103, 111}
_ = string(values) }
