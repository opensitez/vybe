// vybe-test: kotlin/data_class_copying/test_data_class_copying_to_string_equality
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Flag(val enabled: Boolean)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Flag(true)
            val b = a.copy(enabled = false)
            __check((a.toString()).toString(), "Flag(enabled=true)")
            __check((b.toString()).toString(), "Flag(enabled=false)")
            __check((a != b).toString(), "true")
        }
