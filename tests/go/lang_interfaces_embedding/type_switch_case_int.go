// vybe-test: go/lang_interfaces_embedding/type_switch_case_int
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs

package main
import "fmt"
func main() { switch v := any(2).(type) { case int: fmt.Println(v)
default: fmt.Println(0) } }
