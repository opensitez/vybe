// vybe-test: kotlin/primitive_array_apis/test_boolean_array_to_int_mapping
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = booleanArrayOf(true, false, true)
            val mapped = values.map { if (it) 1 else 0 }
            __check((mapped.joinToString(",")).toString(), "1,0,1")
        }
