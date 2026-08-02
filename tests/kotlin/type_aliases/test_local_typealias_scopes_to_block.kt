// vybe-test: kotlin/type_aliases/test_local_typealias_scopes_to_block
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

fun make(): String {
            typealias LocalText = String
            val value: LocalText = "block"
            return value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((make()).toString(), "block")
        }
