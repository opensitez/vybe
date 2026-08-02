// vybe-test: go/regexp_package/regexp_compile_anchor
// origin: languages/go/tests/go/test_regexp_package.rs
// vybe-test-mode: compile

package main
import "regexp"
func main() { _, _ = regexp.Compile(`^start`) }
