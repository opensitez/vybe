// vybe-test: kotlin/kotlin_arrays_creation/test_primitive_to_list_and_join
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = longArrayOf(4L, 8L, 12L)
            __check((values.toList().joinToString(",")).toString(), "4,8,12")
        }
