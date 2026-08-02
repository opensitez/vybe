// vybe-test: go/panic_recover_rules/recover_from_panic_in_map_access
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
m := map[string]int{}
_ = m["missing"]
panic("map") }
func main() { run() }
