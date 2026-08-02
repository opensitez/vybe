// vybe-test: kotlin/kotlin_uri_url/test_uri_resolve_and_normalize
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URI

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = URI("https://example.com/dir/a/b/")
            val child = base.resolve("../c/./index")
            val normalized = child.normalize()
            __check((child.toString()).toString(), "https://example.com/dir/c/index")
            __check((normalized.toString()).toString(), "https://example.com/dir/c/index")
        }
