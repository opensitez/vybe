// vybe-test: go/strconv_package/strconv_append_int_slice
// origin: languages/go/tests/go/test_strconv_package.rs
// vybe-test-mode: compile

package main
import "strconv"
func main() { b := []byte{}
_, _ = strconv.AppendInt(b, 42, 10) }
