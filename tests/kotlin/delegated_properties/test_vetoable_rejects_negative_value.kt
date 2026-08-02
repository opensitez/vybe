// vybe-test: kotlin/delegated_properties/test_vetoable_rejects_negative_value
// origin: languages/kotlin/tests/kotlin/test_delegated_properties.rs

import kotlin.properties.Delegates

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var score by Delegates.vetoable(1) { _, old, new ->
                new >= 0
            }
            score = -3
            val first = score
            score = 7
            val second = score
            __check((first).toString(), "1")
            __check((second).toString(), "7")
        }
