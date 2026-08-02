// vybe-test: go/panic_recover_rules/panic_value_preserved_as_interface
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { v := recover()
fmt.Println(v.(int) + 1) }()
panic(10) }
func main() { run() }
