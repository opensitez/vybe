// vybe-test: go/cover_go_toolchain/build_context_split_path_list
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/build"
func main() { _ = build.Default.SplitPathList("a:b") }
