// vybe-test: go/panic_recover_rules/recover_byte_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
panic(byte(255)) }
func main() { run() }
