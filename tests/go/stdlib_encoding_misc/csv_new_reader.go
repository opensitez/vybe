// vybe-test: go/stdlib_encoding_misc/csv_new_reader
// origin: languages/go/tests/go/test_stdlib_encoding_misc.rs
// vybe-test-mode: compile

package main
import "encoding/csv"
import "strings"
func main() { _ = csv.NewReader(strings.NewReader("a,b")) }
