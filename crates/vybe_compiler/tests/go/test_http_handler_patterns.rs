//! net/http handler patterns: HandlerFunc, ServeMux, StatusText, methods, headers, WithContext.


go_run_cases! {
    http_status_text_ok => (
        "package main; import \"fmt\"; import \"net/http\"; func main() { fmt.Println(http.StatusText(http.StatusOK)) }",
        vec!["OK"]
    ),
    http_status_text_not_found => (
        "package main; import \"fmt\"; import \"net/http\"; func main() { fmt.Println(http.StatusText(http.StatusNotFound)) }",
        vec!["Not Found"]
    ),
    http_method_get_constant_value => (
        "package main; import \"fmt\"; import \"net/http\"; func main() { fmt.Println(http.MethodGet) }",
        vec!["GET"]
    ),
    http_method_post_constant_value => (
        "package main; import \"fmt\"; import \"net/http\"; func main() { fmt.Println(http.MethodPost) }",
        vec!["POST"]
    ),
    http_status_teapot_is_418 => (
        "package main; import \"fmt\"; import \"net/http\"; func main() { fmt.Println(http.StatusTeapot) }",
        vec!["418"]
    ),
}

go_compile_cases! {
    http_handler_func_satisfies_handler => "package main; import \"net/http\"; func main() { var h http.Handler = http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {}) }",
    http_handler_func_write_body => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _, _ = w.Write([]byte(\"ok\")) }) }",
    http_handler_func_read_request_method => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _ = r.Method }) }",
    http_serve_mux_new_zero_value => "package main; import \"net/http\"; func main() { var mux http.ServeMux; _ = mux }",
    http_serve_mux_handle_pattern => "package main; import \"net/http\"; func main() { mux := http.NewServeMux(); mux.Handle(\"/api/\", http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {})) }",
    http_serve_mux_handle_func_sugar => "package main; import \"net/http\"; func main() { mux := http.NewServeMux(); mux.HandleFunc(\"/health\", func(w http.ResponseWriter, r *http.Request) {}) }",
    http_serve_mux_handle_root_pattern => "package main; import \"net/http\"; func main() { mux := http.NewServeMux(); mux.HandleFunc(\"/\", func(w http.ResponseWriter, r *http.Request) {}) }",
    http_serve_mux_handle_nested_path => "package main; import \"net/http\"; func main() { mux := http.NewServeMux(); mux.HandleFunc(\"/v1/users/\", func(w http.ResponseWriter, r *http.Request) {}) }",
    http_default_serve_mux_handle_func => "package main; import \"net/http\"; func main() { http.HandleFunc(\"/ping\", func(w http.ResponseWriter, r *http.Request) {}) }",
    http_default_serve_mux_handle => "package main; import \"net/http\"; func main() { http.Handle(\"/static/\", http.FileServer(http.Dir(\".\"))) }",
    http_serve_mux_serve_http_dispatch => "package main; import \"net/http\"; func main() { mux := http.NewServeMux(); mux.HandleFunc(\"/x\", func(w http.ResponseWriter, r *http.Request) {}); mux.ServeHTTP(nil, nil) }",
    http_response_writer_write_header_before_write => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusCreated); _, _ = w.Write([]byte(\"created\")) }) }",
    http_response_writer_write_header_no_content => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.WriteHeader(http.StatusNoContent) }) }",
    http_response_writer_header_set_content_type => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.Header().Set(\"Content-Type\", \"text/plain\") }) }",
    http_response_writer_header_add_cookie => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.Header().Add(\"Set-Cookie\", \"sid=abc\") }) }",
    http_response_writer_header_get_after_set => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { w.Header().Set(\"X-Trace\", \"1\"); _ = w.Header().Get(\"X-Trace\") }) }",
    http_request_header_get_accept => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com\", nil); req.Header.Set(\"Accept\", \"application/json\"); _ = req.Header.Get(\"Accept\") }",
    http_request_header_add_multiple => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com\", nil); req.Header.Add(\"X-Custom\", \"a\"); req.Header.Add(\"X-Custom\", \"b\"); _ = req.Header[\"X-Custom\"] }",
    http_request_header_del_removes => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com\", nil); req.Header.Set(\"X-Tmp\", \"1\"); req.Header.Del(\"X-Tmp\"); _ = req.Header.Get(\"X-Tmp\") }",
    http_request_with_context_returns_new_request => "package main; import \"context\"; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com\", nil); ctx := context.WithValue(context.Background(), \"k\", 1); _ = req.WithContext(ctx) }",
    http_request_with_context_preserves_url => "package main; import \"context\"; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com/path\", nil); ctx := context.Background(); r2 := req.WithContext(ctx); _ = r2.URL.Path }",
    http_request_context_method => "package main; import \"context\"; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com\", nil); req = req.WithContext(context.Background()); _ = req.Context() }",
    http_method_put_constant => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodPut, \"https://ex.com/item/1\", nil) }",
    http_method_delete_constant => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodDelete, \"https://ex.com/item/1\", nil) }",
    http_method_patch_constant => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodPatch, \"https://ex.com/item/1\", nil) }",
    http_method_head_constant => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodHead, \"https://ex.com/item/1\", nil) }",
    http_method_options_constant => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodOptions, \"https://ex.com\", nil) }",
    http_method_connect_constant => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodConnect, \"https://ex.com\", nil) }",
    http_method_trace_constant => "package main; import \"net/http\"; func main() { _, _ = http.NewRequest(http.MethodTrace, \"https://ex.com\", nil) }",
    http_status_text_bad_request => "package main; import \"net/http\"; func main() { _ = http.StatusText(http.StatusBadRequest) }",
    http_status_text_unauthorized => "package main; import \"net/http\"; func main() { _ = http.StatusText(http.StatusUnauthorized) }",
    http_status_text_forbidden => "package main; import \"net/http\"; func main() { _ = http.StatusText(http.StatusForbidden) }",
    http_status_text_gateway_timeout => "package main; import \"net/http\"; func main() { _ = http.StatusText(http.StatusGatewayTimeout) }",
    http_status_moved_permanently => "package main; import \"net/http\"; func main() { _ = http.StatusMovedPermanently == 301 }",
    http_status_see_other => "package main; import \"net/http\"; func main() { _ = http.StatusSeeOther == 303 }",
    http_status_temporary_redirect => "package main; import \"net/http\"; func main() { _ = http.StatusTemporaryRedirect == 307 }",
    http_status_permanent_redirect => "package main; import \"net/http\"; func main() { _ = http.StatusPermanentRedirect == 308 }",
    http_handler_chain_middleware_pattern => "package main; import \"net/http\"; func main() { wrap := func(next http.Handler) http.Handler { return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { next.ServeHTTP(w, r) }) }; _ = wrap(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {})) }",
    http_serve_mux_handle_method_specific => "package main; import \"net/http\"; func main() { mux := http.NewServeMux(); mux.HandleFunc(\"GET /items\", func(w http.ResponseWriter, r *http.Request) {}) }",
    http_serve_mux_handle_func_post_only => "package main; import \"net/http\"; func main() { mux := http.NewServeMux(); mux.HandleFunc(\"POST /submit\", func(w http.ResponseWriter, r *http.Request) {}) }",
    http_request_form_value_get => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodPost, \"https://ex.com\", nil); _ = req.FormValue(\"field\") }",
    http_request_url_query_get => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com/?q=go\", nil); _ = req.URL.Query().Get(\"q\") }",
    http_request_host_field => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://api.example.com/v1\", nil); _ = req.Host }",
    http_request_proto_major_minor => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com\", nil); _ = req.ProtoMajor; _ = req.ProtoMinor }",
    http_response_writer_flusher_interface => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { if f, ok := w.(http.Flusher); ok { f.Flush() } }) }",
    http_canonical_header_key => "package main; import \"net/http\"; func main() { _ = http.CanonicalHeaderKey(\"content-type\") }",
    http_detect_content_type_sniff => "package main; import \"net/http\"; func main() { _ = http.DetectContentType([]byte(\"<html>\")) }",
    http_error_helper_writes_status => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { http.Error(w, \"bad\", http.StatusBadRequest) }) }",
    http_redirect_helper_location => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { http.Redirect(w, r, \"/login\", http.StatusFound) }) }",
    http_not_found_helper => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { http.NotFound(w, r) }) }",
    http_serve_content_headers => "package main; import \"net/http\"; import \"strings\"; import \"time\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { http.ServeContent(w, r, \"f.txt\", time.Now(), strings.NewReader(\"data\")) }) }",
    http_max_bytes_reader_limit => "package main; import \"net/http\"; import \"strings\"; func main() { req, _ := http.NewRequest(http.MethodPost, \"https://ex.com\", strings.NewReader(\"body\")); _ = http.MaxBytesReader(nil, req.Body, 1024) }",
    http_request_basic_auth_set => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com\", nil); req.SetBasicAuth(\"user\", \"pass\"); _ = req.Header.Get(\"Authorization\") }",
    http_request_user_agent_set => "package main; import \"net/http\"; func main() { req, _ := http.NewRequest(http.MethodGet, \"https://ex.com\", nil); req.Header.Set(\"User-Agent\", \"vybe-test/1.0\"); _ = req.Header.Get(\"User-Agent\") }",
    http_handler_serve_http_method => "package main; import \"net/http\"; type greeter struct{}; func (g greeter) ServeHTTP(w http.ResponseWriter, r *http.Request) {}; func main() { var h http.Handler = greeter{}; h.ServeHTTP(nil, nil) }",
    http_close_notifier_deprecated_compile => "package main; import \"net/http\"; func main() { http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) { _, _ = w.(http.CloseNotifier) }) }",
    http_request_with_cancel_context => "package main; import \"context\"; import \"net/http\"; func main() { ctx, cancel := context.WithCancel(context.Background()); defer cancel(); req, _ := http.NewRequestWithContext(ctx, http.MethodGet, \"https://ex.com\", nil); _ = req.Context().Err() }",
}
