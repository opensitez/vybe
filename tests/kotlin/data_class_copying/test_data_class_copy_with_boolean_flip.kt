// vybe-test: kotlin/data_class_copying/test_data_class_copy_with_boolean_flip
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Switch(val on: Boolean)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Switch(false)
            val b = a.copy(on = true)
            __check((a.on).toString(), "false")
            __check((b.on).toString(), "true")
        }
