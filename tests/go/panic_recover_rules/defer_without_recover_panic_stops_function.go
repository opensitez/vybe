// vybe-test: go/panic_recover_rules/defer_without_recover_panic_stops_function
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer fmt.Println("cleanup")
panic("die")
fmt.Println("skip") }
func main() { defer func() { recover() }()
run()
fmt.Println("ok") }
