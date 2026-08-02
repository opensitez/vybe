// vybe-test: go/flag_parsing_extended/flag_new_flagset_parse_nil_args
// origin: languages/go/tests/go/test_flag_parsing_extended.rs
// vybe-test-mode: compile

package main
import "flag"
func main() { fs := flag.NewFlagSet("tool", flag.ContinueOnError)
_ = fs.Parse(nil) }
