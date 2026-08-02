// vybe-test: go/lang_builtins_control/for_range_map_single_var
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func main() { m := map[string]int{"k":5}
for k := range m { fmt.Println(k)
break } }
