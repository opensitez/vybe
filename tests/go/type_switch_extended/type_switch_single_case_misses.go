// vybe-test: go/type_switch_extended/type_switch_single_case_misses
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
func tag(v interface{}) { switch v.(type) { case int: fmt.Println("hit") } }
func main() { fmt.Println("done") }
