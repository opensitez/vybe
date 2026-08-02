// vybe-test: go/regexp_package/regexp_num_subexp
// origin: languages/go/tests/go/test_regexp_package.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`(a)(b)`)
_ = re.NumSubexp() }
