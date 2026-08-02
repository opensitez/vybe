// vybe-test: go/switch_fallthrough_extended/tagless_switch_no_match_no_default
// origin: languages/go/tests/go/test_switch_fallthrough_extended.rs

package main
import "fmt"
func main() { switch { case false: fmt.Println("no") }
fmt.Println("done") }
