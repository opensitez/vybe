// vybe-test: kotlin/class_delegation/test_delegation_with_nullable_delegate_reference
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Marker { fun tag(): String }

        class MarkerImpl : Marker {
            override fun tag() = "ok"
        }

        class Holder(delegate: Marker) : Marker by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder(MarkerImpl())
            __check((h.tag()).toString(), "ok")
        }
