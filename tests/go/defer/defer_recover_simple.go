// vybe-test: go/defer/defer_recover_simple
// origin: languages/go/tests/go/test_defer.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { defer func() { recover() }()
panic("test")
}
