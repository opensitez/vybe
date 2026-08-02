// vybe-test: kotlin/primitive_array_apis/test_array_reversed_in_place
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = IntArray(4) { it + 1 }
            values.reverse()
            __check((values.joinToString(",")).toString(), "4,3,2,1")
        }
