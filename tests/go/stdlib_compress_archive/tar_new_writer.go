// vybe-test: go/stdlib_compress_archive/tar_new_writer
// origin: languages/go/tests/go/test_stdlib_compress_archive.rs
// vybe-test-mode: compile

package main
import "archive/tar"
import "bytes"
func main() { _ = tar.NewWriter(bytes.NewBuffer(nil)) }
