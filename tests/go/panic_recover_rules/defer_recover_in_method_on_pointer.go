// vybe-test: go/panic_recover_rules/defer_recover_in_method_on_pointer
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
type safe struct{}
func (s *safe) run() { defer func() { fmt.Println(recover() != nil) }()
panic(1) }
func main() { (&safe{}).run() }
