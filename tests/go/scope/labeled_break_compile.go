// vybe-test: go/scope/labeled_break_compile
// origin: languages/go/tests/go/test_scope.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { outer: for i := 0; i < 3; i++ { for j := 0; j < 3; j++ { if j == 1 { break outer }
fmt.Println(i, j)
} } }
