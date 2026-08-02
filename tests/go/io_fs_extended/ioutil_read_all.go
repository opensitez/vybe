// vybe-test: go/io_fs_extended/ioutil_read_all
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io/ioutil"
import "strings"
func main() { _, _ = ioutil.ReadAll(strings.NewReader("hi")) }
