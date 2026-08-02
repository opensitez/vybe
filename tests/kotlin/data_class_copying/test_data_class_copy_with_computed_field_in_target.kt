// vybe-test: kotlin/data_class_copying/test_data_class_copy_with_computed_field_in_target
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Value(val base: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = Value(2)
            val c = v.copy(base = v.base * 3)
            __check((c.base).toString(), "6")
        }
