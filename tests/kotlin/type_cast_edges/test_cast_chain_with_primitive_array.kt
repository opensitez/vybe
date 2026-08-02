// vybe-test: kotlin/type_cast_edges/test_cast_chain_with_primitive_array
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = intArrayOf(1, 2, 3)
            val values = value as? IntArray
            __check((values?.joinToString(",")).toString(), "1,2,3")
        }
