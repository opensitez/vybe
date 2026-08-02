// vybe-test: go/lang_declarations_types/func_not_comparable_compile
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
func f() {}
func g() {}
func main() { _ = f == g }
