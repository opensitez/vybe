// vybe-test: kotlin/primitive_array_apis/test_double_array_join
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = doubleArrayOf(1.5, 2.5)
            __check((values.joinToString(",")).toString(), "1.5,2.5")
            __check((values.sum()).toString(), "4.0")
            __check((values.average()).toString(), "2.0")
        }
