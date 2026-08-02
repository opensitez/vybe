// vybe-test: go/panic_recover_rules/recover_from_panic_in_range
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
for _, v := range []int{1} { if v == 1 { panic("rng") } } }
func main() { run() }
