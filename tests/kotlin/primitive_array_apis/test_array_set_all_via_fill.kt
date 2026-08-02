// vybe-test: kotlin/primitive_array_apis/test_array_set_all_via_fill
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = IntArray(3)
            values.fill(7)
            __check((values.joinToString(",")).toString(), "7,7,7")
        }
