// vybe-test: kotlin/class_delegation/test_delegate_chain_keeps_same_reference
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Marker { fun marker(): String }

        class First : Marker { override fun marker() = "first" }
        class Second(delegate: Marker) : Marker by delegate
        class Third(delegate: Marker) : Marker by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Third(Second(First()))
            __check((t.marker()).toString(), "first")
        }
