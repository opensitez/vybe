// vybe-test: kotlin/primitive_array_apis/test_int_array_get_or_else
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val src = intArrayOf(1, 2)
            __check((src.getOrElse(0) { -1 }).toString(), "1")
            __check((src.getOrElse(4) { -1 }).toString(), "-1")
        }
