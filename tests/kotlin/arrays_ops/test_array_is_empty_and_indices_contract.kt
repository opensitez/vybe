// vybe-test: kotlin/arrays_ops/test_array_is_empty_and_indices_contract
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val none = intArrayOf()
            __check((none.isEmpty()).toString(), "true")
            __check((none.isNotEmpty()).toString(), "false")
            __check((none.indices.joinToString(",")).toString(), "")
        }
