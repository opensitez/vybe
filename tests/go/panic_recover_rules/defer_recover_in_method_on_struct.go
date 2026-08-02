// vybe-test: go/panic_recover_rules/defer_recover_in_method_on_struct
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
type safe struct{}
func (s safe) run() { defer func() { fmt.Println(recover()) }()
panic("m") }
func main() { safe{}.run() }
