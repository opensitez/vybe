// vybe-test: go/json_unmarshal_advanced/json_decoder_token_use_number
// origin: languages/go/tests/go/test_json_unmarshal_advanced.rs
// vybe-test-mode: compile

package main
import "encoding/json"
import "bytes"
func main() { dec := json.NewDecoder(bytes.NewReader([]byte(`{"x":1.5}`)))
dec.UseNumber()
_, _ = dec.Token() }
