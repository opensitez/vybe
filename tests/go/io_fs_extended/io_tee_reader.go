// vybe-test: go/io_fs_extended/io_tee_reader
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io"
import "os"
import "strings"
func main() { _ = io.TeeReader(strings.NewReader("z"), os.Stdout) }
