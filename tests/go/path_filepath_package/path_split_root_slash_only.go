// vybe-test: go/path_filepath_package/path_split_root_slash_only
// origin: languages/go/tests/go/test_path_filepath_package.rs
// vybe-test-mode: compile

package main
import "path"
func main() { _, _ = path.Split("/") }
