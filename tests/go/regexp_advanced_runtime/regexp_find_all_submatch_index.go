// vybe-test: go/regexp_advanced_runtime/regexp_find_all_submatch_index
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`(\d)`)
_ = re.FindAllSubmatchIndex([]byte("a1b2"), -1) }
