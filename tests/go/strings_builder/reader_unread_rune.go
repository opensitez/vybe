// vybe-test: go/strings_builder/reader_unread_rune
// origin: languages/go/tests/go/test_strings_builder.rs
// vybe-test-mode: compile

package main
import "strings"
func main() { r := strings.NewReader("日")
_, _, _ = r.ReadRune()
_ = r.UnreadRune() }
