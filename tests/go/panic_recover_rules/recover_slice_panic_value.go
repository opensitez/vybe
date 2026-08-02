// vybe-test: go/panic_recover_rules/recover_slice_panic_value
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { s := recover().([]int)
fmt.Println(len(s)) }()
panic([]int{1, 2, 3}) }
func main() { run() }
