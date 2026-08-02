// vybe-test: go/regexp_advanced_runtime/regexp_replace_all_byte_slice
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`(\d)`)
_ = re.ReplaceAll([]byte("a1"), []byte("$1")) }
