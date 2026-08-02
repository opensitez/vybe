// vybe-test: go/io_fs_extended/io_read_all
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io"
import "strings"
func main() { _, _ = io.ReadAll(strings.NewReader("ab")) }
