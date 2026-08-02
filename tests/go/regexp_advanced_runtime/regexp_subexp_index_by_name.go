// vybe-test: go/regexp_advanced_runtime/regexp_subexp_index_by_name
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`(?P<id>\d+)`)
_ = re.SubexpIndex("id") }
