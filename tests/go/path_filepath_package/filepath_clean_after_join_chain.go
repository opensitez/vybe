// vybe-test: go/path_filepath_package/filepath_clean_after_join_chain
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path/filepath"
func main() { _ = filepath.Clean(filepath.Join("a", "b", "..", "c")) }
