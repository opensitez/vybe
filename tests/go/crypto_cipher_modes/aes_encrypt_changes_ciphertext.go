// vybe-test: go/crypto_cipher_modes/aes_encrypt_changes_ciphertext
// origin: languages/go/tests/go/test_crypto_cipher_modes.rs

package main
import "fmt"
import "crypto/aes"
func main() { key := []byte{1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16}
block, _ := aes.NewCipher(key)
plain := make([]byte, 16)
for i := range plain { plain[i] = byte(i) }
cipher := make([]byte, 16)
block.Encrypt(cipher, plain)
fmt.Println(cipher[0] != plain[0]) }
