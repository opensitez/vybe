// vybe-test: kotlin/data_classes/test_data_class_with_default_values_preserved_in_copy
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Settings(val enabled: Boolean = true, val retries: Int = 3)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Settings()
            val copy = base.copy(retries = 7)
            __check((base.enabled).toString(), "true")
            __check((base.retries).toString(), "3")
            __check((copy.enabled).toString(), "true")
            __check((copy.retries).toString(), "7")
        }
