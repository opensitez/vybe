// vybe-test: kotlin/kotlin_uri_url/test_url_connection_protocol_metadata_only
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URL

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val url = URL("http://localhost:8080/path?x=1")
            __check((url.protocol).toString(), "http")
            __check((url.host).toString(), "localhost")
            __check((url.port).toString(), "8080")
            __check((url.query).toString(), "x=1")
            __check((url.authority).toString(), "localhost:8080")
            __check((url.file).toString(), "/path?x=1")
        }
