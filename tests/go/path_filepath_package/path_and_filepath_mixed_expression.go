// vybe-test: go/path_filepath_package/path_and_filepath_mixed_expression
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path"
import "path/filepath"
func main() { _ = filepath.Clean(path.Join("src", "main.go")) }
