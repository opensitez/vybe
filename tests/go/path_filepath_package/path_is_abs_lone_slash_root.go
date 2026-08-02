// vybe-test: go/path_filepath_package/path_is_abs_lone_slash_root
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path"
func main() { _ = path.IsAbs("/") }
