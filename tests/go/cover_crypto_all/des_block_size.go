// vybe-test: go/cover_crypto_all/des_block_size
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/des"
func main() { _ = des.BlockSize }
