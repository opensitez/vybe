// vybe-test: kotlin/data_class_copying/test_data_class_copy_with_nullable_property
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class NullableBox(val value: String?)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = NullableBox(null)
            val b = a.copy(value = "x")
            __check((a.value == null).toString(), "true")
            __check((b.value).toString(), "x")
        }
