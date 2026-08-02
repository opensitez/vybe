// vybe-test: kotlin/arrays_ops/test_int_array_to_typed_and_back
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val primitive = intArrayOf(1, 2, 3)
            val boxed = primitive.toTypedArray().toIntArray()
            __check((boxed.joinToString(",")).toString(), "1,2,3")
            __check((primitive.contentEquals(boxed)).toString(), "true")
        }
