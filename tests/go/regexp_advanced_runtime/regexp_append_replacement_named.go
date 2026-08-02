// vybe-test: go/regexp_advanced_runtime/regexp_append_replacement_named
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`(?P<v>\d+)`)
b := re.Expand(nil, []byte("v=$v"), []byte("42"), re.FindSubmatchIndex([]byte("42")))
_ = b }
