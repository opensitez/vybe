// vybe-test: go/flag_parsing_extended/flag_commandline_vs_new_flagset_names
// origin: languages/go/tests/go/test_flag_parsing_extended.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { _ = flag.CommandLine
fs := flag.NewFlagSet("sub", flag.ContinueOnError)
_ = fs }
