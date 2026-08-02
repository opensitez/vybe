// vybe-test: go/cover_encoding_extra/csv_writer_write
// origin: languages/go/tests/go/test_cover_encoding_extra.rs
// vybe-test-mode: compile

package main
import "encoding/csv"
import "bytes"
func main() { w := csv.NewWriter(bytes.NewBuffer(nil))
_ = w.Write([]string{"a", "b"}) }
