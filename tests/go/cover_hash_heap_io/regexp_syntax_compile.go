// vybe-test: go/cover_hash_heap_io/regexp_syntax_compile
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { re, _ := syntax.Parse("a", syntax.Perl)
_ = syntax.Compile(re) }
