// vybe-test: go/panic_recover_rules/panic_message_empty_string
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover() == "") }()
panic("") }
func main() { run() }
