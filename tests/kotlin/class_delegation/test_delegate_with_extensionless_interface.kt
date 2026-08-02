// vybe-test: kotlin/class_delegation/test_delegate_with_extensionless_interface
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Identity { fun id(): String }

        class A : Identity { override fun id() = "A" }
        class Holder(delegate: Identity) : Identity by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder(A())
            __check((h.id()).toString(), "A")
        }
