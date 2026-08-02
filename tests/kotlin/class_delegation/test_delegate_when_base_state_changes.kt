// vybe-test: kotlin/class_delegation/test_delegate_when_base_state_changes
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface MutableCounter { var value: Int }

        class Counter(var value: Int) : MutableCounter

        class Proxy(delegate: MutableCounter) : MutableCounter by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter(1)
            val p = Proxy(c)
            p.value = p.value + 2
            __check((p.value).toString(), "3")
            __check((c.value).toString(), "3")
        }
