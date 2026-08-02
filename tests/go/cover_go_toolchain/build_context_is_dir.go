// vybe-test: go/cover_go_toolchain/build_context_is_dir
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/build"
func main() { _, _ = build.Default.IsDir(".") }
