// vybe-test: kotlin/class_delegation/test_delegate_inheritance_not_allowed_not_used
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Base { fun kind(): String }

        open class Root : Base {
            override fun kind() = "root"
        }

        class Child(delegate: Base) : Base by delegate, Root()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Child(Root()).kind()).toString(), "root")
        }
