// vybe-test: go/stdlib_compress_archive/zlib_new_writer
// origin: languages/go/tests/go/test_stdlib_compress_archive.rs
// vybe-test-mode: compile

package main
import "compress/zlib"
import "bytes"
func main() { _ = zlib.NewWriter(bytes.NewBuffer(nil)) }
