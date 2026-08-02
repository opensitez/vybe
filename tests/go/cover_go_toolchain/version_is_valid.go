// vybe-test: go/cover_go_toolchain/version_is_valid
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/version"
func main() { _ = version.IsValid("go1.21") }
