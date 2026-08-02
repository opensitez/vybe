// vybe-test: kotlin/advanced_features/test_companion_object_method
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

class Counter {
            companion object {
                fun getInit(): Int = 10
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter.getInit()).toString(), "10")
        }
