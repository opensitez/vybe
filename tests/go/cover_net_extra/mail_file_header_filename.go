// vybe-test: go/cover_net_extra/mail_file_header_filename
// origin: languages/go/tests/go/test_cover_net_extra.rs
// vybe-test-mode: compile

package main
import "net/mail"
type FileHeader = mail.FileHeader
func main() { var fh FileHeader
_ = fh.Filename
_ = fh.Header
_ = fh.Size }
