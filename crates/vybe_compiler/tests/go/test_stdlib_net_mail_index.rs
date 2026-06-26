//! net/mail, net/smtp, index/suffixarray, expvar, plugin — breadth compile smokes.

use crate::helpers::*;

go_compile_cases! {
    mail_parse_address => "package main; import \"net/mail\"; func main() { _, _ = mail.ParseAddress(\"Go <go@example.com>\") }",
    mail_parse_date => "package main; import \"net/mail\"; func main() { _, _ = mail.ParseDate(\"Mon, 02 Jan 2006 15:04:05 MST\") }",
    smtp_send_mail => "package main; import \"net/smtp\"; func main() { _ = smtp.SendMail(\"localhost:25\", nil, \"from@example.com\", []string{\"to@example.com\"}, []byte(\"body\")) }",
    smtp_plain_auth => "package main; import \"net/smtp\"; func main() { _ = smtp.PlainAuth(\"\", \"user\", \"pass\", \"localhost\") }",
    suffixarray_new => "package main; import \"index/suffixarray\"; func main() { _ = suffixarray.New([]byte(\"banana\")) }",
    suffixarray_lookup => "package main; import \"index/suffixarray\"; func main() { idx := suffixarray.New([]byte(\"banana\")); _ = idx.Lookup([]byte(\"ana\"), -1) }",
    expvar_publish => "package main; import \"expvar\"; func main() { expvar.Publish(\"hits\", new(expvar.Int)) }",
    expvar_get => "package main; import \"expvar\"; func main() { _ = expvar.Get(\"hits\") }",
    plugin_open => "package main; import \"plugin\"; func main() { _, _ = plugin.Open(\"plugin.so\") }",
}
