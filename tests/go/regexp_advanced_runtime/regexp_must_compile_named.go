// vybe-test: go/regexp_advanced_runtime/regexp_must_compile_named
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { _ = regexp.MustCompile(`(?P<n>\w+)`) }
