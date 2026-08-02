// vybe-test: go/regexp_advanced_runtime/regexp_copy_on_write
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`test`)
_ = re.Copy() }
