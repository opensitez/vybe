// vybe-test: kotlin/kotlin_uri_url/test_uri_parses_components
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URI

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val uri = URI("https://user:pass@example.com:9443/search?q=kotlin&x=1#top")
            __check((uri.scheme).toString(), "https")
            __check((uri.host).toString(), "example.com")
            __check((uri.port).toString(), "9443")
            __check((uri.userInfo).toString(), "user:pass")
            __check((uri.path).toString(), "/search")
            __check((uri.query).toString(), "q=kotlin&x=1")
            __check((uri.fragment).toString(), "top")
            __check((uri.isAbsolute).toString(), "true")
        }
