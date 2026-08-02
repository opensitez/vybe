// vybe-test: go/cover_crypto_all/elliptic_unmarshal
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/elliptic"
func main() { _, _ = elliptic.Unmarshal(elliptic.P256(), []byte{4, 1, 2}) }
