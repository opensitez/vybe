// vybe-test: kotlin/data_classes/test_data_class_with_boolean_and_numeric_fields
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Flagged(val enabled: Boolean, val level: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Flagged(false, 2)
            __check((item.enabled).toString(), "false")
            __check((item.level).toString(), "2")
            val updated = item.copy(enabled = true)
            __check((updated.enabled).toString(), "true")
            __check((updated.level).toString(), "2")
        }
