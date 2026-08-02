// vybe-test: kotlin/kotlin_string_split_replace/test_split_single
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_split_replace.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "a,b,c"
            val parts = s.split(",")
            __check((parts.size).toString(), "3")
            __check((parts[1]).toString(), "b")
        }
