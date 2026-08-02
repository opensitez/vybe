// vybe-test: go/cover_encoding_extra/csv_reader_comment
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/csv"
import "strings"
func main() { r := csv.NewReader(strings.NewReader("#note\na,b"))
r.Comment = '#'
_, _ = r.Read() }
