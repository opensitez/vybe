// vybe-test: go/regexp_package/regexp_find_all_index
// origin: languages/go/tests/go/test_regexp_package.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`x`)
_ = re.FindAllStringIndex("xxy", -1) }
