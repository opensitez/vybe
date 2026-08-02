// vybe-test: kotlin/class_delegation/test_set_delegation_uses_set_contract
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

class ReadOnlySet(delegate: Set<Int>) : Set<Int> by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = ReadOnlySet(setOf(1, 2, 2, 3))
            __check((s.size).toString(), "3")
            __check((s.contains(2)).toString(), "true")
            __check((s.contains(9)).toString(), "false")
        }
