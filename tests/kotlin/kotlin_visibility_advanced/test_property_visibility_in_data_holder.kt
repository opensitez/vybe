// vybe-test: kotlin/kotlin_visibility_advanced/test_property_visibility_in_data_holder
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_advanced.rs

class Repo {
            private val raw = 1
            val exposed get() = raw
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Repo().exposed).toString(), "1")
        }
