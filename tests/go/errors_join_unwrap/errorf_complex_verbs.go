// vybe-test: go/errors_join_unwrap/errorf_complex_verbs
// origin: languages/go/tests/go/test_errors_join_unwrap.rs
// vybe-test-mode: compile

package main
import "fmt"
func main() { _ = fmt.Errorf("%#v %+#v", 1, struct{ X int }{1}) }
