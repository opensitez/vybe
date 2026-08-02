// vybe-test: go/regexp_advanced_runtime/regexp_append_replacement_dollar_one
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
import "bytes"
func main() { re := regexp.MustCompile(`(\w)`)
dst := []byte{}
src := []byte("a")
_ = re.ReplaceAllFunc(src, func(m []byte) []byte { return re.Expand(dst[:0], []byte("[$1]"), src, re.FindIndex(src)) }) }
