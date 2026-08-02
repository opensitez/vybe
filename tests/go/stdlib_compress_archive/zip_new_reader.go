// vybe-test: go/stdlib_compress_archive/zip_new_reader
// origin: languages/go/tests/go/test_stdlib_compress_archive.rs
// vybe-test-mode: compile

package main
import "archive/zip"
import "bytes"
func main() { _, _ = zip.NewReader(bytes.NewReader(nil), 0) }
