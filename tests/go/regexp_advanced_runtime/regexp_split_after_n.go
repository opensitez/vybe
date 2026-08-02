// vybe-test: go/regexp_advanced_runtime/regexp_split_after_n
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`,`)
_ = re.SplitAfterN("a,b,c", 2) }
