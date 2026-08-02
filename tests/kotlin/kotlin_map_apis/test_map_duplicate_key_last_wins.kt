// vybe-test: kotlin/kotlin_map_apis/test_map_duplicate_key_last_wins
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("x" to 1, "x" to 9, "y" to 4)
            __check((map["x"]).toString(), "9")
            __check((map.size).toString(), "2")
        }
