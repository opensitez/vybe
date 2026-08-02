// vybe-test: go/cover_crypto_all/x509_marshal_ec_private_key
// origin: languages/go/tests/go/test_cover_crypto_all.rs
// vybe-test-mode: compile

package main
import "crypto/x509"
import "crypto/ecdsa"
import "crypto/elliptic"
import "crypto/rand"
func main() { key, _ := ecdsa.GenerateKey(elliptic.P256(), rand.Reader)
_, _ = x509.MarshalECPrivateKey(key) }
