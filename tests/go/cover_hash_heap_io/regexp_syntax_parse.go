// vybe-test: go/cover_hash_heap_io/regexp_syntax_parse
// origin: languages/go/tests/go/test_cover_hash_heap_io.rs
// vybe-test-mode: compile

package main
import "regexp/syntax"
func main() { _, _ = syntax.Parse("a+", syntax.Perl) }
