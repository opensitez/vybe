// vybe-test: go/switch_type_tagless/switch_init_statement
// origin: languages/go/tests/go/test_switch_type_tagless.rs

package main
import "fmt"
func main() { switch x := 3; x { case 1: fmt.Println("one") case 3: fmt.Println("three") default: fmt.Println("other") } }
