// vybe-test: go/panic_recover_rules/recover_map_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { m := recover().(map[string]int)
fmt.Println(m["k"]) }()
panic(map[string]int{"k": 4}) }
func main() { run() }
