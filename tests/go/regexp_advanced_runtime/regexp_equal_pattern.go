// vybe-test: go/regexp_advanced_runtime/regexp_equal_pattern
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { a := regexp.MustCompile(`a`)
b := regexp.MustCompile(`a`)
_ = a.String() == b.String() }
