// vybe-test: go/cover_go_toolchain/build_find_only
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/build"
func main() { _ = build.FindOnly }
