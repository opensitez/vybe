// vybe-test: kotlin/data_class_copying/test_data_class_copy_of_copy
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Step(val value: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Step(1)
            val b = a.copy()
            val c = b.copy(value = b.value + 1)
            __check((a.value).toString(), "1")
            __check((b.value).toString(), "1")
            __check((c.value).toString(), "2")
        }
