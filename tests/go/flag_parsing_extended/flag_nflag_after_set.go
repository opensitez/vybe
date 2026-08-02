// vybe-test: go/flag_parsing_extended/flag_nflag_after_set
// origin: languages/go/tests/go/test_flag_parsing_extended.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { _ = flag.Bool("v", false, "")
_ = flag.Set("v", "true")
_ = flag.NFlag() }
