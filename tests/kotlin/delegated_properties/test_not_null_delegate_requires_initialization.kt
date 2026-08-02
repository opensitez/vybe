// vybe-test: kotlin/delegated_properties/test_not_null_delegate_requires_initialization
// origin: languages/kotlin/tests/kotlin/test_delegated_properties.rs

import kotlin.properties.Delegates

        class Box {
            var value: Int by Delegates.notNull()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box()
            val out = try {
                box.value
                "ok"
            } catch (e: IllegalStateException) {
                "not-set"
            }
            __check((out).toString(), "not-set")
            box.value = 11
            __check((box.value).toString(), "11")
        }
