// vybe-test: kotlin/kotlin_uri_url/test_uri_file_scheme_roundtrip
// origin: languages/kotlin/tests/kotlin/test_kotlin_uri_url.rs

import java.net.URI

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
            val root = java.lang.System.getProperty("java.io.tmpdir")
            val uri = URI("file", null, root, 0, "/tmp.log", null, null)
            __p((uri.scheme).toString())
            __p((uri.path).toString())
            __p((uri.isAbsolute).toString())
            __p((uri.toString().startsWith("file:")).toString())
        
__check("file\n/tmp.log\ntrue\ntrue")
}
