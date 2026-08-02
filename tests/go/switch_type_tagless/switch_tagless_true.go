// vybe-test: go/switch_type_tagless/switch_tagless_true
// origin: languages/go/tests/go/test_switch_type_tagless.rs

package main
import "fmt"
func main() { x := 5
switch { case x < 3: fmt.Println("low") case x < 10: fmt.Println("mid") default: fmt.Println("high") } }
