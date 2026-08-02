// vybe-test: go/stdlib_compress_archive/flate_new_reader
// origin: languages/go/tests/go/test_stdlib_compress_archive.rs
// vybe-test-mode: compile

package main
import "compress/flate"
import "bytes"
func main() { _ = flate.NewReader(bytes.NewReader(nil)) }
