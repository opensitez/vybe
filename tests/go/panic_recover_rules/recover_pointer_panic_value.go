// vybe-test: go/panic_recover_rules/recover_pointer_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { p := recover().(*int)
fmt.Println(*p) }()
n := 6
panic(&n) }
func main() { run() }
