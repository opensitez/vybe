// vybe-test: kotlin/primitive_array_apis/test_array_transform_to_set
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = intArrayOf(1, 1, 2, 2, 3)
            val set = values.toMutableSet()
            __check((set.joinToString(",")).toString(), "1,2,3")
            __check((set.size).toString(), "3")
        }
