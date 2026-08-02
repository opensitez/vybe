// vybe-test: kotlin/literals/test_collection_from_literal_values
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf(1, 2, 3).joinToString(",")).toString(), "1,2,3")
            __check((listOf("a", "b", "c").size).toString(), "3")
            __check((arrayOf(1.0, 2.5).joinToString("|")).toString(), "1.0|2.5")
            __check((intArrayOf(1, 2, 3).size).toString(), "3")
        }
