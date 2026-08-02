// vybe-test: go/strings_unicode_extra/string_switch_runtime
// origin: languages/go/tests/go/test_strings_unicode_extra.rs

package main
import "fmt"
func main() { text := "go"
switch text { case "go": fmt.Println(1)
default: fmt.Println(0) } }
