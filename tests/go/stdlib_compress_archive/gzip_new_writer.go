// vybe-test: go/stdlib_compress_archive/gzip_new_writer
// origin: languages/go/tests/go/test_stdlib_compress_archive.rs
// vybe-test-mode: compile

package main
import "compress/gzip"
import "bytes"
func main() { _ = gzip.NewWriter(bytes.NewBuffer(nil)) }
