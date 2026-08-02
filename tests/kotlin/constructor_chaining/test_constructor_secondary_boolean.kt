// vybe-test: kotlin/constructor_chaining/test_constructor_secondary_boolean
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Flag {
            val enabled: Boolean
            constructor(enabled: Boolean) { this.enabled = enabled }
            constructor() : this(false)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Flag().enabled).toString(), "false")
            __check((Flag(true).enabled).toString(), "true")
        }
