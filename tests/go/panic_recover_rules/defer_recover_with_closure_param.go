// vybe-test: go/panic_recover_rules/defer_recover_with_closure_param
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func(label string) { if recover() != nil { fmt.Println(label) } }("ok")
panic("x") }
func main() { run() }
