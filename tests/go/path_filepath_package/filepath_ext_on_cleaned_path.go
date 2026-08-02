// vybe-test: go/path_filepath_package/filepath_ext_on_cleaned_path
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Ext(filepath.Clean("dir/file.GO")) }
