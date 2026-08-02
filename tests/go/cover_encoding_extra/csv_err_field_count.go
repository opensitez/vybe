// vybe-test: go/cover_encoding_extra/csv_err_field_count
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/csv"
func main() { _ = csv.ErrFieldCount }
