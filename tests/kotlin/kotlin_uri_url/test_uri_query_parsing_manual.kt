// vybe-test: kotlin/kotlin_uri_url/test_uri_query_parsing_manual
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URI

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val uri = URI("https://example.com/a?left=1&right=2")
            val query = uri.query
            val parts = query.split("&").joinToString("|") { it }
            __check((parts).toString(), "left=1|right=2")
            __check((query.contains("left=1")).toString(), "true")
            __check((query.contains("right=2")).toString(), "true")
        }
