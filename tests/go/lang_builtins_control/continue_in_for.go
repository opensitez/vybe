// vybe-test: go/lang_builtins_control/continue_in_for
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func main() { for i := 0; i < 3; i++ { if i == 1 { continue }
fmt.Println(i) } }
