// vybe-test: kotlin/kotlin_lazy_delegates/test_delegates_vetoable_rejects_invalid_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

import kotlin.properties.Delegates

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value by Delegates.vetoable(1) { _, _, newValue ->
                newValue >= 0
            }
            value = 3
            __check((value).toString(), "3")
            value = -10
            __check((value).toString(), "3")
        }
