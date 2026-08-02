// vybe-test: go/cover_debug_formats/gosym_decoding_error
// origin: languages/go/tests/go/test_cover_debug_formats.rs
// vybe-test-mode: compile

package main
import "debug/gosym"
func main() { var e gosym.DecodingError
_ = e.Error() }
