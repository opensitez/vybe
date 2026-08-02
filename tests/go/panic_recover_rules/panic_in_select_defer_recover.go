// vybe-test: go/panic_recover_rules/panic_in_select_defer_recover
// origin: languages/go/tests/go/test_panic_recover_rules.rs

package main
import "fmt"
func run() { defer func() { fmt.Println(recover()) }()
ch := make(chan int, 1)
ch <- 1
select { case <-ch: panic("sel") } }
func main() { run() }
