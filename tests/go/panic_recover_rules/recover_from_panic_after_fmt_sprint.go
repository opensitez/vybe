// vybe-test: go/panic_recover_rules/recover_from_panic_after_fmt_sprint
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
_ = fmt.Sprint(1)
panic("fmt") }
func main() { run() }
