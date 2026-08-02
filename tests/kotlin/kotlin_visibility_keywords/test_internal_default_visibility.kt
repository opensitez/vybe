// vybe-test: kotlin/kotlin_visibility_keywords/test_internal_default_visibility
// origin: languages/kotlin/tests/kotlin/test_kotlin_visibility_keywords.rs

internal const val scope = "module"

        class Counter {
            internal var value = 0
            fun bump(): String {
                value = value + 1
                return scope + value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter()
            __check((c.bump()).toString(), "module1")
            __check((c.bump()).toString(), "module2")
        }
