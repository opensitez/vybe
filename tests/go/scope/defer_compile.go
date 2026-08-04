// vybe-test: go/scope/defer_compile
// origin: languages/go/tests/go/test_scope.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { defer fmt.Println("deferred")
fmt.Println("first")
}
