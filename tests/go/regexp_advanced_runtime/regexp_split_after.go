// vybe-test: go/regexp_advanced_runtime/regexp_split_after
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`\s`)
_ = re.SplitAfter("a b c", -1) }
