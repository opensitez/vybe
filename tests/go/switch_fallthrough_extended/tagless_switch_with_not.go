// vybe-test: go/switch_fallthrough_extended/tagless_switch_with_not
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { ready := false
switch { case !ready: fmt.Println("wait")
default: fmt.Println("go") } }
