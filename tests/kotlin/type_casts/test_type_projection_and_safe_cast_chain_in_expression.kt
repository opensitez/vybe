// vybe-test: kotlin/type_casts/test_type_projection_and_safe_cast_chain_in_expression
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = arrayOf(1, 2, 3)
            val values = (value as? IntArray) ?: intArrayOf(9, 8)
            __check((values.joinToString(",")).toString(), "9,8")
        }
