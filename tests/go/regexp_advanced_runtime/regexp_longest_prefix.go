// vybe-test: go/regexp_advanced_runtime/regexp_longest_prefix
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`foo`)
_ = re.Longest() }
