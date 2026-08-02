// vybe-test: go/cover_go_toolchain/version_compare
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/version"
func main() { _ = version.Compare("go1.21", "go1.22") }
