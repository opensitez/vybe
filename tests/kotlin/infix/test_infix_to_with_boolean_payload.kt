// vybe-test: kotlin/infix/test_infix_to_with_boolean_payload
// origin: languages/kotlin/tests/kotlin/test_infix.rs

data class PairBool(val key: String, val enabled: Boolean)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = "feature" to true
            __check((item.first).toString(), "feature")
            __check((item.second).toString(), "true")
            val toggled = item.first to (item.second && false)
            __check((toggled.second).toString(), "false")
        }
