// vybe-test: go/strconv_extended/strconv_can_backquote_unicode
// origin: languages/go/tests/go/test_strconv_extended.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { _ = strconv.CanBackquote("日本語") }
