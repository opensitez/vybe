// vybe-test: go/lang_declarations_types/labeled_for_continue
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func main() { L: for i := 0; i < 3; i++ { if i == 1 { continue L }
fmt.Println(i) } }
