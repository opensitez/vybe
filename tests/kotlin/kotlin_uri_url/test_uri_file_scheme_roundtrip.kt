// vybe-test: kotlin/kotlin_uri_url/test_uri_file_scheme_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URI

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val root = java.lang.System.getProperty("java.io.tmpdir")
            val uri = URI("file", null, root, 0, "/tmp.log", null, null)
            __check((uri.scheme).toString(), "file")
            __check((uri.path).toString(), "/tmp.log")
            __check((uri.isAbsolute).toString(), "true")
            __check((uri.toString().startsWith("file:")).toString(), "true")
        }
