// vybe-test: go/unicode_utf8/utf8_string_range_rune_values
// origin: languages/go/tests/go/test_unicode_utf8.rs

package main
import "fmt"
func main() { total := 0
for _, r := range "日本語" { total += int(r) }
fmt.Println(total) }
