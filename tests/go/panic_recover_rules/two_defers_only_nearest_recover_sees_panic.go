// vybe-test: go/panic_recover_rules/two_defers_only_nearest_recover_sees_panic
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() == nil) }()
defer func() { fmt.Println(recover() != nil) }()
panic("boom") }
func main() { run() }
