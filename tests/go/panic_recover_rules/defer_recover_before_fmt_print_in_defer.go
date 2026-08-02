// vybe-test: go/panic_recover_rules/defer_recover_before_fmt_print_in_defer
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { if r := recover(); r != nil { fmt.Println("got") } else { fmt.Println("none") } }()
panic("p") }
func main() { run() }
