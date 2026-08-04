// vybe-test: go/scope/goto_compile
// origin: languages/go/tests/go/test_scope.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { goto done
fmt.Println("skip")
done: fmt.Println("done")
}
