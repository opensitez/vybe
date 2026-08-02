// vybe-test: kotlin/kotlin_uri_url/test_uri_relativize_behavior
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URI

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = URI("https://example.org/root/index")
            val target = URI("https://example.org/root/docs/page")
            val rel = base.relativize(target)
            __check((rel.toString()).toString(), "docs/page")
            __check((base.resolve(rel) == target).toString(), "true")
        }
