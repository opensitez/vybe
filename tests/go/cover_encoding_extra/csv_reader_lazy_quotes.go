// vybe-test: go/cover_encoding_extra/csv_reader_lazy_quotes
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/csv"
import "strings"
func main() { r := csv.NewReader(strings.NewReader(`"a",b`))
r.LazyQuotes = true
_, _ = r.Read() }
