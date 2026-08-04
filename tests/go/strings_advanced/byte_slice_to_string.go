// vybe-test: go/strings_advanced/byte_slice_to_string
// origin: languages/go/tests/go/test_strings_advanced.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { b := []byte{97, 98, 99}
s := string(b)
_ = s }
