// vybe-test: go/path_filepath_package/path_ext_leading_dotfile_empty
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path"
func main() { _ = path.Ext(".profile") }
