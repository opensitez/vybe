// vybe-test: go/panic_recover_rules/recover_complex_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() != nil) }()
panic(complex(1, 2)) }
func main() { run() }
