//! net/http subpackages — one API per test.

use crate::helpers::*;

go_compile_cases! {
    httptest_new_server => "package main; import \"net/http/httptest\"; func main() { s := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {})); s.Close() }",
    httptest_new_request => "package main; import \"net/http/httptest\"; func main() { _ = httptest.NewRequest(\"GET\", \"/\", nil) }",
    httptest_response_recorder => "package main; import \"net/http/httptest\"; func main() { _ = httptest.NewRecorder() }",
    httputil_reverse_proxy => "package main; import \"net/http/httputil\"; func main() { _ = httputil.NewSingleHostReverseProxy(nil) }",
    httputil_dump_request => "package main; import \"net/http/httputil\"; func main() { _, _ = httputil.DumpRequest(nil, false) }",
    httputil_dump_response => "package main; import \"net/http/httputil\"; func main() { _, _ = httputil.DumpResponse(nil, false) }",
    httptrace_context_client_trace => "package main; import \"net/http/httptrace\"; func main() { _ = httptrace.WithClientTrace(nil, &httptrace.ClientTrace{}) }",
    pprof_handler => "package main; import \"net/http/pprof\"; func main() { _ = pprof.Handler(\"goroutine\") }",
    pprof_lookup => "package main; import \"runtime/pprof\"; func main() { _, _ = pprof.Lookup(\"goroutine\") }",
    pprof_start_cpu => "package main; import \"runtime/pprof\"; import \"os\"; func main() { _ = pprof.StartCPUProfile(os.Stdout) }",
    pprof_write_heap => "package main; import \"runtime/pprof\"; import \"os\"; func main() { _ = pprof.WriteHeapProfile(os.Stdout) }",
    cgi_handler => "package main; import \"net/http/cgi\"; func main() { _ = cgi.Handler{} }",
}
