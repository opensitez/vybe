// vybe-test: kotlin/data_classes/test_data_class_default_values_preserve_when_copying_subset
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Flag(val enabled: Boolean = true, val count: Int = 1)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Flag()
            val copied = base.copy(count = 7)
            __check((base.enabled).toString(), "true")
            __check((base.count).toString(), "1")
            __check((copied.enabled).toString(), "true")
            __check((copied.count).toString(), "7")
        }
