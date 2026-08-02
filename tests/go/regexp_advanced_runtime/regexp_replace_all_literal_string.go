// vybe-test: go/regexp_advanced_runtime/regexp_replace_all_literal_string
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`a+`)
_ = re.ReplaceAllLiteralString("aaa", "b") }
