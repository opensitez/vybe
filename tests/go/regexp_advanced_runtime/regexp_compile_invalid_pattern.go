// vybe-test: go/regexp_advanced_runtime/regexp_compile_invalid_pattern
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { _, err := regexp.Compile(`(`)
_ = err }
