// vybe-test: go/switch_fallthrough_extended/switch_in_for_loop
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { for i := 0; i < 2; i++ { switch i { case 0: fmt.Println("z")
case 1: fmt.Println("o")
} } }
