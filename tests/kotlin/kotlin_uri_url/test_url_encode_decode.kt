// vybe-test: kotlin/kotlin_uri_url/test_url_encode_decode
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URLEncoder
        import java.net.URLDecoder

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val encoded = URLEncoder.encode("a b/c", "UTF-8")
            val decoded = URLDecoder.decode(encoded, "UTF-8")
            __check((encoded).toString(), "a+b%2Fc")
            __check((decoded).toString(), "a b/c")
        }
