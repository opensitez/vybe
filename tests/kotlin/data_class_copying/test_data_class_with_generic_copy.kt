// vybe-test: kotlin/data_class_copying/test_data_class_with_generic_copy
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Holder<T>(val value: T)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Holder("x")
            val b = a.copy(value = "y")
            __check((a.value).toString(), "x")
            __check((b.value).toString(), "y")
        }
