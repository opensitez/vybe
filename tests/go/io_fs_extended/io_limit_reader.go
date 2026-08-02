// vybe-test: go/io_fs_extended/io_limit_reader
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { _ = io.LimitReader(strings.NewReader("abcd"), 2) }
