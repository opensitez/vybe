// vybe-test: kotlin/delegated_properties/test_vetoable_allows_true_transition
// origin: languages/kotlin/tests/kotlin/test_delegated_properties.rs

import kotlin.properties.Delegates

        var total by Delegates.vetoable(0) { _, _, new -> new >= 0 }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            total = 5
            total = 6
            __check((total).toString(), "6")
        }
