// vybe-test: kotlin/primitive_array_apis/test_array_sorting
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = intArrayOf(4, 1, 3, 2)
            values.sort()
            __check((values.joinToString(",")).toString(), "1,2,3,4")
        }
