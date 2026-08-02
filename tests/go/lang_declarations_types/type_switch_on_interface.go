// vybe-test: go/lang_declarations_types/type_switch_on_interface
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func main() { switch any("x").(type) { case string: fmt.Println("s")
default: fmt.Println("d") } }
