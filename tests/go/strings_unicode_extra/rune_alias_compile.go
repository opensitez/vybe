// vybe-test: go/strings_unicode_extra/rune_alias_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
type myRune rune
func main() { var r myRune = 'a'
_ = r }
