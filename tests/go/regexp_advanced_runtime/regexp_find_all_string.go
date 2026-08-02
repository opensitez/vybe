// vybe-test: go/regexp_advanced_runtime/regexp_find_all_string
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`\d+`)
_ = re.FindAllString("a1b22", -1) }
