// vybe-test: kotlin/kotlin_uri_url/test_uri_to_url_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URI

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val uri = URI("https://example.org/resource")
            val url = uri.toURL()
            __check((url.protocol).toString(), "https")
            __check((url.host).toString(), "example.org")
            __check((url.path).toString(), "/resource")
            __check((url.toURI() == uri).toString(), "true")
        }
