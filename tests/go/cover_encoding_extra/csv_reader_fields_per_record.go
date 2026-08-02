// vybe-test: go/cover_encoding_extra/csv_reader_fields_per_record
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/csv"
import "strings"
func main() { r := csv.NewReader(strings.NewReader("a,b"))
r.FieldsPerRecord = 2
_, _ = r.Read() }
