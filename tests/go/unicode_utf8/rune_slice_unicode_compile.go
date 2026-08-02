// vybe-test: go/unicode_utf8/rune_slice_unicode_compile
// origin: languages/go/tests/go/test_unicode_utf8.rs
// vybe-test-mode: compile

package main
func main() { rs := []rune("café")
_ = rs[4] }
