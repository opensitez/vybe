// vybe-test: go/stdlib_encoding_misc/csv_new_writer
// origin: languages/go/tests/go/test_stdlib_encoding_misc.rs
// vybe-test-mode: compile

package main
import "encoding/csv"
import "bytes"
func main() { _ = csv.NewWriter(bytes.NewBuffer(nil)) }
