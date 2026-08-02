// vybe-test: kotlin/collections/test_array_plus_operator_is_not_mutating_source
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = arrayOf(1, 2)
            val right = arrayOf(3, 4)
            val joined = left + right
            __check((joined.joinToString(",")).toString(), "1,2,3,4")
            left[0] = 9
            __check((joined[0]).toString(), "1")
        }
