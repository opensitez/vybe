// vybe-test: kotlin/strings/test_split_and_reconstruct
// origin: languages/kotlin/tests/kotlin/test_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parts = "a,b,c,d".split(",")
            __check((parts.size).toString(), "4")
            __check((parts[0]).toString(), "a")
            __check((parts[3]).toString(), "d")
            __check((parts.joinToString("|")).toString(), "a|b|c|d")
        }
