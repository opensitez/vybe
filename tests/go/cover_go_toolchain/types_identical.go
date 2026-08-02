// vybe-test: go/cover_go_toolchain/types_identical
// origin: languages/go/tests/go/test_cover_go_toolchain.rs
// vybe-test-mode: compile

package main
import "go/types"
func main() { _ = types.Identical(types.Typ[types.Int], types.Typ[types.Int]) }
