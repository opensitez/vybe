// vybe-test: go/regexp_advanced_runtime/regexp_find_reader
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
import "strings"
func main() { re := regexp.MustCompile(`(\d+)`)
_, _ = re.FindReaderSubmatchIndex(strings.NewReader("n42")) }
