// vybe-test: kotlin/advanced_features/test_advanced_nested_object_expression_with_state
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun makeCounter() = object {
            var value: Int = 0

            fun inc(): Int {
                value += 1
                return value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = makeCounter()
            __check((c.inc()).toString(), "1")
            __check((c.inc()).toString(), "2")
        }
