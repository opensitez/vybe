//! net/rpc, net/rpc/jsonrpc, net/mail, net/smtp — breadth compile smokes.

go_compile_cases! {
    // net/rpc — Server
    rpc_new_server => "package main; import \"net/rpc\"; func main() { _ = rpc.NewServer() }",
    rpc_default_server => "package main; import \"net/rpc\"; func main() { _ = rpc.DefaultServer }",
    rpc_server_register => "package main; import \"net/rpc\"; type Args struct { A, B int }; type Arith int; func (t *Arith) Add(args *Args, reply *int) error { return nil }; func main() { s := rpc.NewServer(); s.Register(new(Arith)) }",
    rpc_server_register_name => "package main; import \"net/rpc\"; type Args struct { A, B int }; type Arith int; func (t *Arith) Mul(args *Args, reply *int) error { return nil }; func main() { s := rpc.NewServer(); s.RegisterName(\"Math\", new(Arith)) }",
    rpc_register_package_level => "package main; import \"net/rpc\"; type Args struct { A, B int }; type Arith int; func (t *Arith) Sub(args *Args, reply *int) error { return nil }; func main() { rpc.Register(new(Arith)) }",
    rpc_register_name_package_level => "package main; import \"net/rpc\"; type Args struct { A, B int }; type Arith int; func (t *Arith) Div(args *Args, reply *int) error { return nil }; func main() { rpc.RegisterName(\"Calc\", new(Arith)) }",
    rpc_server_handle_http => "package main; import \"net/rpc\"; func main() { s := rpc.NewServer(); s.HandleHTTP() }",
    rpc_server_accept => "package main; import \"net\"; import \"net/rpc\"; func main() { s := rpc.NewServer(); ln, _ := net.Listen(\"tcp\", \"127.0.0.1:0\"); defer ln.Close(); go s.Accept(ln) }",

    // net/rpc — Client
    rpc_dial => "package main; import \"net/rpc\"; func main() { _, _ = rpc.Dial(\"tcp\", \"127.0.0.1:9999\") }",
    rpc_dial_http => "package main; import \"net/rpc\"; func main() { _, _ = rpc.DialHTTP(\"tcp\", \"127.0.0.1:9999\") }",
    rpc_new_client => "package main; import \"net\"; import \"net/rpc\"; func main() { conn, _ := net.Dial(\"tcp\", \"127.0.0.1:9999\"); if conn != nil { defer conn.Close(); _, _ = rpc.NewClient(conn) } }",
    rpc_client_call => "package main; import \"net/rpc\"; type Args struct { A, B int }; func main() { c, _ := rpc.Dial(\"tcp\", \"127.0.0.1:9999\"); if c != nil { defer c.Close(); _ = c.Call(\"Arith.Add\", &Args{1, 2}, new(int)) } }",
    rpc_client_go => "package main; import \"net/rpc\"; type Args struct { A, B int }; func main() { c, _ := rpc.Dial(\"tcp\", \"127.0.0.1:9999\"); if c != nil { defer c.Close(); _ = c.Go(\"Arith.Add\", &Args{1, 2}, new(int), make(chan *rpc.Call, 1)) } }",
    rpc_client_close => "package main; import \"net/rpc\"; func main() { c, _ := rpc.Dial(\"tcp\", \"127.0.0.1:9999\"); if c != nil { _ = c.Close() } }",

    // net/rpc — types
    rpc_call_type => "package main; import \"net/rpc\"; type Call = rpc.Call; func main() { var call Call; _ = call.ServiceMethod; _ = call.Reply; _ = call.Error }",
    rpc_request_type => "package main; import \"net/rpc\"; type Request = rpc.Request; func main() { var req Request; _ = req.ServiceMethod; _ = req.Seq; _ = req.Args }",
    rpc_response_type => "package main; import \"net/rpc\"; type Response = rpc.Response; func main() { var resp Response; _ = resp.ServiceMethod; _ = resp.Seq; _ = resp.Error }",
    rpc_server_error => "package main; import \"net/rpc\"; func main() { _ = rpc.ServerError }",

    // net/rpc/jsonrpc
    jsonrpc_new_client => "package main; import \"net\"; import \"net/rpc/jsonrpc\"; func main() { conn, _ := net.Dial(\"tcp\", \"127.0.0.1:9999\"); if conn != nil { defer conn.Close(); _ = jsonrpc.NewClient(conn) } }",
    jsonrpc_new_client_codec => "package main; import \"net\"; import \"net/rpc/jsonrpc\"; func main() { conn, _ := net.Dial(\"tcp\", \"127.0.0.1:9999\"); if conn != nil { defer conn.Close(); _ = jsonrpc.NewClientCodec(conn) } }",
    jsonrpc_new_server_codec => "package main; import \"net\"; import \"net/rpc/jsonrpc\"; func main() { conn, _ := net.Dial(\"tcp\", \"127.0.0.1:9999\"); if conn != nil { defer conn.Close(); _ = jsonrpc.NewServerCodec(conn) } }",
    jsonrpc_serve_conn => "package main; import \"net\"; import \"net/rpc/jsonrpc\"; func main() { ln, _ := net.Listen(\"tcp\", \"127.0.0.1:0\"); defer ln.Close(); conn, _ := ln.Accept(); if conn != nil { jsonrpc.ServeConn(conn) } }",

    // net/mail — distinct from ParseAddress/ParseDate smokes elsewhere
    mail_parse_address_list => "package main; import \"net/mail\"; func main() { _, _ = mail.ParseAddressList(\"Alice <a@example.com>, Bob <b@example.com>\") }",
    mail_read_message => "package main; import \"net/mail\"; import \"strings\"; func main() { _, _ = mail.ReadMessage(strings.NewReader(\"Subject: hi\\r\\n\\r\\nbody\")) }",
    mail_message_header => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"To: you@example.com\\r\\n\\r\\n\")); _ = msg.Header }",
    mail_header_get => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"Subject: hi\\r\\n\\r\\n\")); _ = msg.Header.Get(\"Subject\") }",
    mail_header_set => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"\\r\\n\")); msg.Header.Set(\"Subject\", \"Hello\") }",
    mail_header_add => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"\\r\\n\")); msg.Header.Add(\"Received\", \"from localhost\") }",
    mail_header_date => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"Date: Mon, 02 Jan 2006 15:04:05 MST\\r\\n\\r\\n\")); _, _ = msg.Header.Date() }",
    mail_header_subject => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"Subject: test\\r\\n\\r\\n\")); _ = msg.Header.Subject() }",
    mail_header_message_id => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"Message-ID: <id@example.com>\\r\\n\\r\\n\")); _ = msg.Header.MessageID() }",
    mail_address_string => "package main; import \"net/mail\"; func main() { addr, _ := mail.ParseAddress(\"Go Team <go@example.com>\"); _ = addr.String() }",
    mail_address_address_field => "package main; import \"net/mail\"; func main() { addr, _ := mail.ParseAddress(\"dev@example.com\"); _ = addr.Address; _ = addr.Name }",
    mail_message_body => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"\\r\\nhello\")); _ = msg.Body }",
    mail_message_write_to => "package main; import \"bytes\"; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"Subject: x\\r\\n\\r\\n\")); var buf bytes.Buffer; _, _ = msg.WriteTo(&buf) }",
    mail_word_encoder => "package main; import \"net/mail\"; type Encoder = mail.WordEncoder; func main() { var enc Encoder; _ = enc.Encode(\"hello world\") }",
    mail_word_decoder => "package main; import \"net/mail\"; type Decoder = mail.WordDecoder; func main() { var dec Decoder; _, _ = dec.Decode(\"hello\") }",

    // net/smtp — distinct from SendMail/PlainAuth smokes elsewhere
    smtp_dial => "package main; import \"net/smtp\"; func main() { _, _ = smtp.Dial(\"localhost:25\") }",
    smtp_client_auth => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Auth(smtp.PlainAuth(\"\", \"u\", \"p\", \"localhost\")) } }",
    smtp_client_close => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { _ = c.Close() } }",
    smtp_client_data => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _, _ = c.Data() } }",
    smtp_client_extension => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _, _ = c.Extension(\"STARTTLS\") } }",
    smtp_client_hello => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Hello(\"localhost\") } }",
    smtp_client_mail => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Mail(\"from@example.com\") } }",
    smtp_client_quit => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Quit() } }",
    smtp_client_rcpt => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Rcpt(\"to@example.com\") } }",
    smtp_client_reset => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Reset() } }",
    smtp_client_start_tls => "package main; import \"crypto/tls\"; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.StartTLS(&tls.Config{ServerName: \"localhost\"}) } }",
    smtp_client_verify => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Verify(\"user@example.com\") } }",
    smtp_cram_md5_auth => "package main; import \"net/smtp\"; func main() { _ = smtp.CRAMMD5Auth(\"user\", \"secret\") }",
    smtp_client_text_field => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Text } }",

    // net/rpc — ServeRequest and codec types
    rpc_default_server_handle_http => "package main; import \"net/rpc\"; func main() { rpc.DefaultServer.HandleHTTP() }",
    rpc_serve_request_type => "package main; import \"net/rpc\"; type Server = rpc.Server; func main() { var s Server; _ = s.ServeRequest }",
    rpc_client_codec_field => "package main; import \"net/rpc\"; func main() { c, _ := rpc.Dial(\"tcp\", \"127.0.0.1:9999\"); if c != nil { defer c.Close(); _ = c.Codec } }",
    rpc_server_codec_type => "package main; import \"net/rpc\"; type ServerCodec = rpc.ServerCodec; func main() { var sc ServerCodec; _ = sc }",
    rpc_client_codec_type => "package main; import \"net/rpc\"; type ClientCodec = rpc.ClientCodec; func main() { var cc ClientCodec; _ = cc }",
    mail_message_attachments => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"Content-Type: multipart/mixed\\r\\n\\r\\n\")); _ = msg.Attachments }",
    mail_file_header_filename => "package main; import \"net/mail\"; type FileHeader = mail.FileHeader; func main() { var fh FileHeader; _ = fh.Filename; _ = fh.Header; _ = fh.Size }",
    mail_header_map_len => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"Subject: hi\\r\\n\\r\\n\")); _ = len(msg.Header) }",
    smtp_data_writer => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); w, _ := c.Data(); if w != nil { _, _ = w.Write([]byte(\"hello\")) } } }",
    smtp_auth_cram_md5 => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); _ = c.Auth(smtp.CRAMMD5Auth(\"user\", \"secret\")) } }",
    jsonrpc_client_call => "package main; import \"net/rpc/jsonrpc\"; import \"net\"; func main() { conn, _ := net.Dial(\"tcp\", \"127.0.0.1:9999\"); if conn != nil { defer conn.Close(); c := jsonrpc.NewClient(conn); _ = c.Call(\"Arith.Add\", struct{ A, B int }{1, 2}, new(int)) } }",
    rpc_err_shutdown => "package main; import \"net/rpc\"; func main() { _ = rpc.ErrShutdown }",
    mail_header_values => "package main; import \"net/mail\"; import \"strings\"; func main() { msg, _ := mail.ReadMessage(strings.NewReader(\"To: a@example.com\\r\\n\\r\\n\")); _ = msg.Header.Values(\"To\") }",
    smtp_local_name => "package main; import \"net/smtp\"; func main() { c, _ := smtp.Dial(\"localhost:25\"); if c != nil { defer c.Close(); c.LocalName = \"localhost\"; _ = c.LocalName } }",
}
