// vybe-test: kotlin/kotlin_string_split_replace/test_split_limit
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_split_replace.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "a,b,c,d"
            val parts = s.split(",", limit = 2)
            __check((parts.size).toString(), "2")
            __check((parts[1]).toString(), "b,c,d")
        }
