// vybe-test: go/regexp_advanced_runtime/regexp_expand_dst_buffer
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { re := regexp.MustCompile(`(\d+)`)
_ = re.Expand([]byte{}, []byte("n=$1"), []byte("7"), re.FindSubmatchIndex([]byte("7"))) }
