// vybe-test: go/io_fs_extended/ioutil_write_file_compile
// origin: languages/go/tests/go/test_io_fs_extended.rs
// vybe-test-mode: compile

package main
import "io/ioutil"
func main() { _ = ioutil.WriteFile("out.txt", []byte("x"), 0644) }
