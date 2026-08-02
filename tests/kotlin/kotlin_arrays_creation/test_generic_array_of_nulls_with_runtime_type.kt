// vybe-test: kotlin/kotlin_arrays_creation/test_generic_array_of_nulls_with_runtime_type
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = arrayOfNulls<String>(3)
            values[0] = "a"
            values[1] = "b"
            values[2] = "c"
            __check((values.filterNotNull().joinToString(",")).toString(), "a,b,c")
        }
