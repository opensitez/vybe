// vybe-test: go/io_fs_extended/io_copy_compile
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { _, _ = io.Copy(io.Discard, strings.NewReader("a")) }
