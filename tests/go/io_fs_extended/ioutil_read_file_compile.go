// vybe-test: go/io_fs_extended/ioutil_read_file_compile
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io/ioutil"
func main() { _, _ = ioutil.ReadFile("/dev/null") }
