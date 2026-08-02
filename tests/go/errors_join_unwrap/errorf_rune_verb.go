// vybe-test: go/errors_join_unwrap/errorf_rune_verb
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { _ = fmt.Errorf("char %c", 'A') }
