// vybe-test: go/lang_builtins_control/for_range_string_index
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func main() { for i := range "ab" { if i == 1 { fmt.Println(i)
break } } }
