// vybe-test: go/strings_unicode_extra/unicode_rune_in_switch_compile
// origin: languages/go/tests/go/test_strings_unicode_extra.rs
// vybe-test-mode: compile

package main
func main() { switch 'λ' { case 'λ': _ = 1 } }
