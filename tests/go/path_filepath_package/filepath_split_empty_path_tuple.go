// vybe-test: go/path_filepath_package/filepath_split_empty_path_tuple
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _, _ = filepath.Split("") }
