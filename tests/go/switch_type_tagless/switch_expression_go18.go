// vybe-test: go/switch_type_tagless/switch_expression_go18
// origin: languages/go/tests/go/test_switch_type_tagless.rs

package main
import "fmt"
func main() { x := 2
switch x { case 1, 2: fmt.Println("pair") default: fmt.Println("other") } }
