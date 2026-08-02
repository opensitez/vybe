// vybe-test: go/errors_join_unwrap/errors_join_four_constituents
// origin: languages/go/tests/go/test_errors_join_unwrap.rs

package main
import "fmt"
import "errors"
func main() { err := errors.Join(errors.New("a"), errors.New("b"), errors.New("c"), errors.New("d"))
parts := 0
for _, ch := range err.Error() { if ch == '\n' { parts++ } }
fmt.Println(parts) }
