// vybe-test: go/lang_interfaces_embedding/method_on_defined_string_type
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs
// vybe-test-mode: compile

package main
type MyString string
func (s MyString) Len() int { return len(s) }
func main() { _ = MyString("a").Len() }
