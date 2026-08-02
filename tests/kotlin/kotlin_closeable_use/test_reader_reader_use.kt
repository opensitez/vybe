// vybe-test: kotlin/kotlin_closeable_use/test_reader_reader_use
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.BufferedReader
        import java.io.StringReader

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = "alpha\nbeta"
            val reader = StringReader(text)
            val out = BufferedReader(reader).use { br ->
                br.readLine() + "|" + br.readLine()
            }
            __check((out).toString(), "alpha|beta")
        }
