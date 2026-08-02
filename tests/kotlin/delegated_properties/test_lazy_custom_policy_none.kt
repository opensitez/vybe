// vybe-test: kotlin/delegated_properties/test_lazy_custom_policy_none
// origin: languages/kotlin/tests/kotlin/test_delegated_properties.rs

import kotlin.LazyThreadSafetyMode

        class Holder {
            var invoked = 0
            val value by lazy(LazyThreadSafetyMode.NONE) {
                invoked += 1
                "done"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder()
            __check((h.value).toString(), "done")
            __check((h.value).toString(), "done")
            __check((h.invoked).toString(), "1")
        }
