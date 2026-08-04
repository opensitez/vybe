// vybe-test: kotlin/kotlin_uri_url/test_url_connection_protocol_metadata_only
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URL

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val url = URL("http://localhost:8080/path?x=1")
            __p((url.protocol).toString())
            __p((url.host).toString())
            __p((url.port).toString())
            __p((url.query).toString())
            __p((url.authority).toString())
            __p((url.file).toString())
        
__check("http\nlocalhost\n8080\nx=1\nlocalhost:8080\n/path?x=1")
}
