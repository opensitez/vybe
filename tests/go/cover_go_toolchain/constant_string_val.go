// vybe-test: go/cover_go_toolchain/constant_string_val
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/constant"
func main() { _ = constant.String(constant.Make(1)) }
