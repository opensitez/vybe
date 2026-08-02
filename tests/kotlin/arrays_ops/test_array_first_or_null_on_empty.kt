// vybe-test: kotlin/arrays_ops/test_array_first_or_null_on_empty
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val none = intArrayOf()
            __check((none.firstOrNull() ?: "missing").toString(), "missing")
            __check((none.lastOrNull() ?: "missing").toString(), "missing")
        }
