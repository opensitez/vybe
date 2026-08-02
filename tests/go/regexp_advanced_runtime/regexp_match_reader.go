// vybe-test: go/regexp_advanced_runtime/regexp_match_reader
// origin: languages/go/tests/go/test_regexp_advanced_runtime.rs
// vybe-test-mode: compile

package main
import "regexp"
import "strings"
func main() { re := regexp.MustCompile(`go`)
_, _ = re.MatchReader(strings.NewReader("gopher")) }
