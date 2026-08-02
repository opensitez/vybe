// vybe-test: go/io_fs_extended/ioutil_nop_closer
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io/ioutil"
func main() { _ = ioutil.NopCloser(nil) }
